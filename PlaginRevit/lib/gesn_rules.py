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


def _get_sheet_path(zip_file, sheet_name):
    rels_ns = {
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }

    with zip_file.open("xl/workbook.xml") as data:
        wb_tree = ElementTree.parse(data)
    sheet_id = None
    for sheet in wb_tree.getroot().iterfind("s:sheets/s:sheet", rels_ns):
        if sheet.get("name") == sheet_name:
            sheet_id = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            break
    if not sheet_id:
        return None

    with zip_file.open("xl/_rels/workbook.xml.rels") as data:
        rels_tree = ElementTree.parse(data)
    for rel in rels_tree.getroot().iterfind("r:Relationship", rels_ns):
        if rel.get("Id") == sheet_id:
            target = rel.get("Target")
            if target.startswith("/"):
                target = target[1:]
            return "xl/" + target if not target.startswith("xl/") else target
    return None


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


def _load_sheet_as_rows(excel_path, sheet_name):
    """Простое чтение XLSX без внешних зависимостей."""
    with zipfile.ZipFile(excel_path, "r") as zf:
        sheet_path = _get_sheet_path(zf, sheet_name)
        if not sheet_path:
            raise ValueError(u"Лист '{0}' не найден в файле правил".format(sheet_name))

        shared_strings = _load_shared_strings(zf)
        return _read_sheet_rows(zf, sheet_path, shared_strings)


def load_rules_from_excel(path=None, sheet_name=None):
    """Загрузка правил из Excel.

    Возвращает список GesnRule. Пропускает пустые строки и строки без кода ГЭСН.
    """
    excel_path = path or config.EXCEL_PATH
    sheet = sheet_name or config.EXCEL_SHEET_NAME

    if not os.path.exists(excel_path):
        raise IOError(u"Файл правил не найден: {0}".format(excel_path))

    rows = _load_sheet_as_rows(excel_path, sheet)
    if not rows:
        return []

    header = rows[0]
    header_map = {name: idx for idx, name in enumerate(header)}

    def get_cell(row, name):
        idx = header_map.get(name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    rules = []
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
