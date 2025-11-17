# -*- coding: utf-8 -*-
"""Работа с правилами подбора ГЭСН из Excel."""
import os
import re
import zipfile
from collections import namedtuple
from xml.etree import ElementTree

from . import config

GesnRule = namedtuple(
    "GesnRule",
    [
        "family",
        "type_name",
        "thickness_mm",
        "height_min_mm",
        "height_max_mm",
        "reinforcement",
        "brick_size",
        "gesn_code",
        "unit_raw",
        "multiplier",
        "volume_param",
    ],
)


def _as_text(value):
    if value is None:
        return None
    try:
        return unicode(value)  # type: ignore[name-defined]
    except Exception:
        try:
            return str(value)
        except Exception:
            return None


def _as_float(value):
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def _normalize_bool_text(value):
    text = (_as_text(value) or u"").strip().lower()
    if not text:
        return u""
    return u"да" if text in {u"да", u"yes", u"1", u"true", u"истина"} else u"нет"


def _load_shared_strings(zip_file):
    """Чтение sharedStrings.xml в словарь индексов."""
    try:
        with zip_file.open("xl/sharedStrings.xml") as data:
            tree = ElementTree.parse(data)
    except KeyError:
        return []

    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [
        u"".join(node.itertext())
        for node in tree.getroot().iterfind("s:si", namespace)
    ]


def _column_index(cell_ref):
    match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
    if not match:
        return 0
    col_letters = match.group(1)
    index = 0
    for ch in col_letters:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def _normalize_sheet_name(name):
    return (_as_text(name) or u"").replace(" ", "").replace("_", "").lower()


def _get_sheet_entries(zip_file):
    """Получение списка всех листов с путями к файлам worksheet."""

    rels_ns = {
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }

    with zip_file.open("xl/workbook.xml") as data:
        wb_tree = ElementTree.parse(data)

    try:
        with zip_file.open("xl/_rels/workbook.xml.rels") as data:
            rels_tree = ElementTree.parse(data)
    except KeyError:
        rels_tree = None

    rels_map = {}
    if rels_tree:
        for rel in rels_tree.getroot().iterfind("r:Relationship", rels_ns):
            rels_map[rel.get("Id")] = rel.get("Target")

    sheets_raw = []
    for sheet in wb_tree.getroot().iterfind("s:sheets/s:sheet", rels_ns):
        sheets_raw.append(
            (
                sheet.get("name"),
                sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"),
            )
        )

    def _resolve_target(sheet_record):
        _, rel_id = sheet_record
        if not rel_id:
            return None
        target = rels_map.get(rel_id)
        if not target:
            return None
        if target.startswith("/"):
            target = target[1:]
        return "xl/" + target if not target.startswith("xl/") else target

    entries = []
    for sheet_record in sheets_raw:
        resolved = _resolve_target(sheet_record)
        if resolved:
            entries.append((sheet_record[0], resolved))

    available_names = [name for name, _ in sheets_raw]
    return entries, available_names


def _order_sheets(entries, preferred_name):
    """Возвращает листы в порядке: сначала совпадающие с предпочтительным именем."""

    if not preferred_name:
        return entries

    normalized = _normalize_sheet_name(preferred_name)
    matched = []
    other = []
    for entry in entries:
        if entry[0] == preferred_name or _normalize_sheet_name(entry[0]) == normalized:
            matched.append(entry)
        else:
            other.append(entry)

    return matched + other


def _read_sheet_rows(zip_file, sheet_path, shared_strings):
    """Чтение строк листа xlsx в виде списков значений."""
    with zip_file.open(sheet_path) as data:
        tree = ElementTree.parse(data)

    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = []
    for row in tree.getroot().iterfind("s:sheetData/s:row", namespace):
        cells = []
        for cell in row.iterfind("s:c", namespace):
            idx = _column_index(cell.get("r", "A1"))
            while len(cells) <= idx:
                cells.append(None)

            cell_type = cell.get("t")
            value_node = cell.find("s:v", namespace)
            if value_node is None:
                cells[idx] = None
                continue

            raw_value = value_node.text or u""
            if cell_type == "s":
                try:
                    cells[idx] = shared_strings[int(raw_value)]
                except Exception:
                    cells[idx] = raw_value
            else:
                cells[idx] = raw_value
        rows.append(cells)
    return rows


def _load_all_sheets_as_rows(excel_path, sheet_name):
    """Чтение всех листов XLSX без внешних зависимостей."""

    with zipfile.ZipFile(excel_path, "r") as zf:
        entries, available = _get_sheet_entries(zf)
        if not entries:
            raise ValueError(
                u"В файле правил нет доступных листов. Найдены имена: {0}".format(
                    u", ".join(available)
                )
            )

        ordered_entries = _order_sheets(entries, sheet_name)
        shared_strings = _load_shared_strings(zf)

        sheets_rows = []
        for name, path in ordered_entries:
            try:
                rows = _read_sheet_rows(zf, path, shared_strings)
            except Exception:
                continue
            sheets_rows.append((name, rows))

        return sheets_rows


def load_rules_from_excel(path=None, sheet_name=None):
    """Загрузка правил из Excel.

    Возвращает список GesnRule. Пропускает пустые строки и строки без кода ГЭСН.
    """
    excel_path = path or config.EXCEL_PATH
    sheet = sheet_name or config.EXCEL_SHEET_NAME

    if not os.path.exists(excel_path):
        raise IOError(u"Файл правил не найден: {0}".format(excel_path))

    sheets_rows = _load_all_sheets_as_rows(excel_path, sheet)
    rules = []

    for sheet_name, rows in sheets_rows:
        if not rows:
            continue

        header = rows[0]
        header_map = {name: idx for idx, name in enumerate(header)}
        if u"ГЭСН_код" not in header_map:
            continue

        def get_cell(row, name):
            idx = header_map.get(name)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        for row in rows[1:]:
            gesn_code = _as_text(get_cell(row, u"ГЭСН_код"))
            if not gesn_code:
                continue

            rule = GesnRule(
                family=_as_text(get_cell(row, u"Family")) or u"",
                type_name=_as_text(get_cell(row, u"TypeName")) or u"",
                thickness_mm=_as_float(get_cell(row, u"Thickness_mm")) or 0.0,
                height_min_mm=_as_float(get_cell(row, u"Height_min_mm")) or 0.0,
                height_max_mm=_as_float(get_cell(row, u"Height_max_mm")) or 0.0,
                reinforcement=_normalize_bool_text(get_cell(row, u"Армирование")),
                brick_size=_as_text(get_cell(row, u"Размеры кирпича")) or u"",
                gesn_code=gesn_code,
                unit_raw=_as_text(get_cell(row, u"ЕдИзм")) or u"",
                multiplier=_as_float(get_cell(row, u"Кратность")) or 1.0,
                volume_param=_as_text(get_cell(row, u"Параметр_объёма")) or u"",
            )
            rules.append(rule)

    return rules
