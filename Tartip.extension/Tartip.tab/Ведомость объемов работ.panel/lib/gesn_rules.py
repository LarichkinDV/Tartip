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
        "stage",
        "reinforcement",
        "brick_size",
        "gesn_code",
        "unit_raw",
        "multiplier",
        "volume_param",
        "height_conditions",
        "volume_conditions",
        "height_label",
        "volume_label",
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


def _normalize_stage_value(value):
    text = (_as_text(value) or u"").replace(u"\xa0", u" ")
    norm = text.strip().lower()
    aliases = {
        u"reconstruction": u"реконструкция",
        u"reconstruction stage": u"реконструкция",
        u"new construction": u"новая конструкция",
        u"newconstruction": u"новая конструкция",
        u"existing": u"существующая",
        u"phase created": u"",
    }
    return aliases.get(norm, norm)


def _parse_conditions(raw_value, default_operator=None):
    """Парсинг строковых условий вида ">1000&<=2000" в список (op, number)."""

    text = _as_text(raw_value)
    if text is None:
        return [], u""

    cleaned = text.replace(" ", "").replace("*", "&").replace(",", ".")
    parts = re.split(r"[&;|]+", cleaned)

    conditions = []
    labels = []
    for part in parts:
        if not part:
            continue
        match = re.match(r"(<=|>=|<|>|=)?([-+]?[0-9]*\.?[0-9]+)", part)
        if not match:
            continue
        op = match.group(1) or default_operator or "="
        num = _as_float(match.group(2))
        if num is None:
            continue
        conditions.append((op, num))
        labels.append(u"{0}{1}".format(op, num))

    # Fallback: если не удалось разобрать (например, текст с числом внутри),
    # пробуем вытащить все числа и опциональные операторы из исходного текста.
    if not conditions:
        for match in re.finditer(r"(<=|>=|<|>|=)?\\s*([-+]?[0-9]*\\.?[0-9]+)", text):
            op = match.group(1) or default_operator or "="
            num = _as_float(match.group(2))
            if num is None:
                continue
            conditions.append((op, num))
            labels.append(u"{0}{1}".format(op, num))

    return conditions, u" & ".join(labels)


def _build_height_conditions(raw_min, raw_max):
    """Собирает условия высоты из отдельных ячеек минимального/максимального значения."""

    conditions = []
    labels = []

    min_conditions, min_label = _parse_conditions(raw_min, default_operator=">=")
    max_conditions, max_label = _parse_conditions(raw_max, default_operator="<=")

    conditions.extend(min_conditions)
    conditions.extend(max_conditions)

    if min_label:
        labels.append(min_label)
    if max_label:
        labels.append(max_label)

    return conditions, u" & ".join(labels)


def _first_number(raw_value, fallback=None):
    """Пытается извлечь первое число для обратной совместимости."""

    direct = _as_float(raw_value)
    if direct is not None:
        return direct

    conditions, _ = _parse_conditions(raw_value)
    if conditions:
        return conditions[0][1]

    return fallback


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
    """Получение списка всех листов с путями к файлам worksheet.

    При отсутствии рабочей таблицы связей пытаемся использовать стандартные
    пути вида ``xl/worksheets/sheet{N}.xml``.
    """

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
    zip_names = set(zip_file.namelist())
    for index, sheet_record in enumerate(sheets_raw, start=1):
        resolved = _resolve_target(sheet_record)
        if resolved:
            entries.append((sheet_record[0], resolved))
            continue

        # Если отношения не заданы, пробуем стандартный путь sheet{N}.xml
        fallback_path = "xl/worksheets/sheet{0}.xml".format(index)
        if fallback_path in zip_names:
            entries.append((sheet_record[0], fallback_path))

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
        header_map = {}
        for idx, raw_name in enumerate(header):
            key = (_as_text(raw_name) or u"").strip()
            if key:
                header_map[key] = idx
        if u"Шифр ГЭСН" not in header_map:
            continue

        def get_cell(row, name):
            idx = header_map.get(name)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        volume_cond_column = None
        for candidate in [
            u"Объем_условие",
            u"Условие объема",
            u"Объем",
            u"Volume_condition",
            u"VolumeRange",
        ]:
            if candidate in header_map:
                volume_cond_column = candidate
                break

        has_stage_column = u"Стадия" in header_map
        thickness_headers = [u"Width", u"Ширина", u"Толщина"]

        for row in rows[1:]:
            gesn_code = _as_text(get_cell(row, u"Шифр ГЭСН"))
            if not gesn_code:
                continue

            raw_height_value = get_cell(row, u"Неприсоединенная высота")
            raw_height_text = (_as_text(raw_height_value) or u"").strip()
            if raw_height_value is None or not raw_height_text:
                # Пустое значение высоты в Excel означает отсутствие ограничения.
                height_conditions = []
                height_label = u""
                height_min_mm = None
                height_max_mm = None
            else:
                height_conditions, height_label = _build_height_conditions(
                    None, raw_height_value
                )
                height_min_mm = 0.0
                height_max_mm = _first_number(raw_height_value, 0.0) or 0.0

            if has_stage_column:
                stage_raw = get_cell(row, u"Стадия")
                stage = _normalize_stage_value(stage_raw)
            else:
                stage = u""

            thickness_mm = None
            for header in thickness_headers:
                if header not in header_map:
                    continue
                candidate = _as_float(get_cell(row, header))
                if candidate is None:
                    continue
                thickness_mm = candidate
                break

            volume_conditions = []
            volume_label = u""
            if volume_cond_column:
                volume_conditions, volume_label = _parse_conditions(
                    get_cell(row, volume_cond_column)
                )

            rule = GesnRule(
                family=(_as_text(get_cell(row, u"Семейство")) or u"").strip(),
                type_name=(_as_text(get_cell(row, u"Тип")) or u"").strip(),
                thickness_mm=thickness_mm,
                height_min_mm=height_min_mm,
                height_max_mm=height_max_mm,
                stage=stage,
                reinforcement=_normalize_bool_text(get_cell(row, u"Армирование")),
                brick_size=(
                    (_as_text(get_cell(row, u"Размеры кладочного материала")) or u"")
                    .strip()
                    .lower()
                ),
                gesn_code=gesn_code,
                unit_raw=_as_text(get_cell(row, u"Единица измерения")) or u"",
                multiplier=_as_float(get_cell(row, u"Кратность единицы измерения"))
                or 1.0,
                volume_param=(
                    (_as_text(get_cell(row, u"Параметр_объёма")) or u"").strip()
                    or getattr(config, "DEFAULT_VOLUME_PARAM", u"Площадь")
                ),
                height_conditions=height_conditions,
                volume_conditions=volume_conditions,
                height_label=height_label,
                volume_label=volume_label,
            )
            rules.append(rule)

    return rules


def collect_column_values_from_excel(path=None, sheet_name=None, columns=None):
    """Возвращает уникальные значения указанных колонок из Excel."""

    if not columns:
        return {}

    excel_path = path or config.EXCEL_PATH
    sheet = sheet_name or config.EXCEL_SHEET_NAME

    if not os.path.exists(excel_path):
        raise IOError(u"Файл правил не найден: {0}".format(excel_path))

    sheets_rows = _load_all_sheets_as_rows(excel_path, sheet)
    result = {column: set() for column in columns}

    for _, rows in sheets_rows:
        if not rows:
            continue

        header = rows[0]
        header_map = {}
        for idx, raw_name in enumerate(header):
            key = (_as_text(raw_name) or u"").strip()
            if key:
                header_map[key] = idx

        column_indices = {
            column: header_map.get(column)
            for column in result.keys()
        }

        for row in rows[1:]:
            for column, values in result.items():
                idx = column_indices.get(column)
                if idx is None or idx >= len(row):
                    continue
                text = (_as_text(row[idx]) or u"").strip()
                if text:
                    values.add(text)

    return {column: values for column, values in result.items() if values}


def load_rules_from_db():
    """Заглушка для последующей реализации загрузки правил из БД."""

    raise NotImplementedError(u"Загрузка правил из БД пока не реализована")
