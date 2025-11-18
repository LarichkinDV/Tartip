# -*- coding: utf-8 -*-
"""Определение ГЭСН для стен по таблице соответствий."""
import os
import sys

from pyrevit import revit, DB, forms, script

# Добавляем путь к общим модулям расширения
THIS_DIR = os.path.dirname(__file__)
BASE_DIR = os.path.dirname(THIS_DIR)
LIB_DIR = os.path.join(BASE_DIR, "lib")
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from lib.gesn_rules import load_rules_from_excel  # noqa: E402
from lib import config  # noqa: E402

# Имена используемых параметров
PARAM_REINFORCEMENT = u"Армирование"
PARAM_BRICK_SIZE = u"Размеры кирпича"
# Параметры для вывода результата: сначала приоритетный, затем совместимый резервный
PARAM_GESN_OUTPUT = [u"ACBD_ГЭСН", u"Шифр ГЭСН"]

# Внутренние системные параметры
PARAM_UNCONNECTED_HEIGHT = DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM
PARAM_WIDTH = DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM

# Перевод единиц
FEET_TO_MM = 304.8

# Соответствие имени параметра объёма нужному BuiltInParameter
VOLUME_PARAMS = {
    u"Площадь": DB.BuiltInParameter.HOST_AREA_COMPUTED,
    u"Объем": DB.BuiltInParameter.HOST_VOLUME_COMPUTED,
}


def _h(value):
    """Экранирование HTML для безопасного вывода."""

    if value is None:
        return u""
    text = _t(value)
    return (
        text.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def _t(value):
    try:
        return unicode(value)  # type: ignore[name-defined]
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _normalize_bool_text(value):
    text = (_t(value) or u"").strip().lower()
    if not text:
        return u""
    return u"да" if text in {u"да", u"yes", u"true", u"1", u"истина"} else u"нет"


def _get_parameter_value(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return None
    if param.StorageType == DB.StorageType.Integer:
        try:
            return bool(param.AsInteger())
        except Exception:
            return None
    if param.StorageType == DB.StorageType.String:
        try:
            return param.AsString()
        except Exception:
            return None
    try:
        return param.AsValueString()
    except Exception:
        return None


def _get_height_mm(wall):
    param = wall.get_Parameter(PARAM_UNCONNECTED_HEIGHT)
    return (param.AsDouble() if param else 0.0) * FEET_TO_MM


def _get_thickness_mm(wall_type):
    param = wall_type.get_Parameter(PARAM_WIDTH)
    return (param.AsDouble() if param else 0.0) * FEET_TO_MM


def _get_type(wall):
    try:
        return wall.WallType
    except Exception:
        try:
            return revit.doc.GetElement(wall.GetTypeId())
        except Exception:
            return None


def _prepare_rules():
    return load_rules_from_excel()


def _get_writable_param(element, names):
    """Возвращает первый доступный параметр из списка имен."""

    for name in names:
        param = element.LookupParameter(name)
        if param and not param.IsReadOnly:
            return param
    return None


def _match_rules(wall, rules, thickness_mm, height_mm, reinforcement_text, brick_size):
    wall_type = _get_type(wall)
    # Защищаем получение имен типа и семейства от отсутствующих свойств
    family_name = _t(getattr(wall_type, "FamilyName", u""))
    type_name = _t(getattr(wall_type, "Name", u""))

    matched = []
    for rule in rules:
        if rule.family and rule.family != family_name:
            continue
        if rule.type_name and rule.type_name != type_name:
            continue
        if abs(rule.thickness_mm - thickness_mm) > config.THICKNESS_TOLERANCE_MM:
            continue
        if not (rule.height_min_mm <= height_mm <= rule.height_max_mm):
            continue
        if rule.reinforcement and rule.reinforcement != reinforcement_text:
            continue
        if rule.brick_size and rule.brick_size != brick_size:
            continue
        matched.append(rule)
    return matched


def _get_volume_value(wall, rule):
    target_param = VOLUME_PARAMS.get(rule.volume_param)
    if not target_param:
        return None, None

    param = wall.get_Parameter(target_param)
    if not param:
        return None, None

    raw_value = param.AsDouble()
    if target_param == DB.BuiltInParameter.HOST_AREA_COMPUTED:
        units_label = u"м2"
        metric_value = DB.UnitUtils.ConvertFromInternalUnits(raw_value, DB.UnitTypeId.SquareMeters)
    else:
        units_label = u"м3"
        metric_value = DB.UnitUtils.ConvertFromInternalUnits(raw_value, DB.UnitTypeId.CubicMeters)

    return metric_value, units_label


def _calc_code_fragment(rule, volume_value):
    multiplier = rule.multiplier or 1.0
    volume_param = rule.volume_param or u""
    return u"{0}[{1}/{2}]".format(
        rule.gesn_code,
        volume_param,
        int(multiplier) if multiplier.is_integer() else multiplier,
    )


def _format_rule_result(rule, volume_value, unit_label):
    """Формирование строки для отчёта с объёмом."""

    multiplier = rule.multiplier or 1.0
    normalized = volume_value / multiplier if multiplier else volume_value
    volume_param = rule.volume_param or u""
    parts = [
        u"{0} — {1}: {2:.3f} {3}".format(
            rule.gesn_code,
            volume_param,
            volume_value,
            unit_label or u"",
        )
    ]
    parts.append(u"кратность: {0}".format(int(multiplier) if multiplier.is_integer() else multiplier))
    parts.append(u"объём для ГЭСН: {0:.3f}".format(normalized))
    return u"; ".join(parts)


def _process_wall(wall, rules):
    """Обработка одной стены и запись результата/причины в параметр."""

    entry = {
        "id": getattr(getattr(wall, "Id", None), "IntegerValue", None),
        "cat": _t(getattr(getattr(wall, "Category", None), "Name", u"")),
        "type": None,
        "family": None,
        "message": None,
        "matched": False,
    }

    target_param = _get_writable_param(wall, PARAM_GESN_OUTPUT)

    if not target_param:
        # Даже причину записать некуда
        entry["message"] = u"Нет доступного параметра для записи"
        return False, False, entry

    wall_type = _get_type(wall)
    if wall_type is None:
        reason = u"Не удалось определить тип стены"
        entry["message"] = reason
        return target_param.Set(reason), False, entry

    entry["type"] = _t(getattr(wall_type, "Name", u""))
    entry["family"] = _t(getattr(wall_type, "FamilyName", u""))

    thickness_mm = _get_thickness_mm(wall_type)
    height_mm = _get_height_mm(wall)
    reinforcement_text = _normalize_bool_text(_get_parameter_value(wall, PARAM_REINFORCEMENT))
    brick_size = _t(_get_parameter_value(wall, PARAM_BRICK_SIZE))

    matched_rules = _match_rules(wall, rules, thickness_mm, height_mm, reinforcement_text, brick_size)
    if not matched_rules:
        reason = u"Нет подходящей записи в БД"
        entry["message"] = reason
        return target_param.Set(reason), False, entry

    fragments_for_param = []
    fragments_for_report = []
    last_volume_issue = None
    for rule in matched_rules:
        volume_value, unit_label = _get_volume_value(wall, rule)
        if volume_value is None:
            last_volume_issue = u"Не найден параметр объёма: {0}".format(rule.volume_param or u"?")
            continue
        fragments_for_param.append(_calc_code_fragment(rule, volume_value))
        fragments_for_report.append(_format_rule_result(rule, volume_value, unit_label))

    if not fragments_for_param:
        reason = last_volume_issue or u"Не удалось вычислить объём"
        entry["message"] = reason
        return target_param.Set(reason), False, entry

    entry["matched"] = True
    entry["message"] = u"; ".join(fragments_for_report)
    return target_param.Set(u"; ".join(fragments_for_param)), True, entry


def _collect_walls():
    uidoc = revit.uidoc
    doc = revit.doc
    selection_ids = list(uidoc.Selection.GetElementIds()) if uidoc else []

    if selection_ids:
        elements = [doc.GetElement(eid) for eid in selection_ids]
        return [el for el in elements if getattr(el, "Category", None) and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Walls)]

    collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
    return list(collector)


def main():
    out = script.get_output()
    try:
        rules = _prepare_rules()
    except Exception as exc:
        forms.alert(u"Не удалось загрузить таблицу правил: {0}".format(exc), exitscript=True)
        return

    walls = _collect_walls()
    if not walls:
        forms.alert(u"В модели не найдены стены для обработки", exitscript=True)
        return

    processed = 0
    updated = 0
    matched = 0
    entries = []

    with revit.Transaction(u"ТАРТИП: определить ГЭСН"):
        for wall in walls:
            ok, has_match, entry = _process_wall(wall, rules)
            entries.append(entry)
            if ok:
                processed += 1
                if has_match:
                    matched += 1
                    updated += 1
                elif config.CLEAR_CODE_WHEN_MISS and not entry.get("matched"):
                    param = _get_writable_param(wall, PARAM_GESN_OUTPUT)
                    if param:
                        param.Set(u"")
            else:
                out.print_html(u"Не удалось обновить стену: {0}".format(_t(wall)))

    not_matched = processed - matched
    forms.alert(
        u"Обработано стен: {0}\nОбновлено ГЭСН: {1}\nБез подходящей записи: {2}".format(
            processed, updated, not_matched
        )
    )

    # Выводим перечень элементов с найденными работами или причинами отсутствия
    try:
        out.clear()
    except Exception:
        pass

    rows = []
    for entry in entries:
        link = (
            out.linkify(DB.ElementId(entry["id"]), u"{}".format(entry["id"]))
            if entry.get("id") is not None
            else u""
        )
        rows.append(
            [
                link,
                _h(entry.get("cat") or u""),
                _h(entry.get("family") or u""),
                _h(entry.get("type") or u""),
                _h(entry.get("message") or u""),
            ]
        )

    if rows:
        header = [u"ID", u"Категория", u"Семейство", u"Тип", u"Результат"]
        css = (
            u"<style>table.acbd{border-collapse:collapse;width:100%;margin:6px 0;}"
            u"table.acbd th,table.acbd td{border:1px solid #d0d0d0;padding:4px 6px;}"
            u"table.acbd thead th{background:#f6f6f6;position:sticky;top:0;}</style>"
        )
        table_html = [css, u"<table class='acbd'>", u"<thead><tr>"]
        for h in header:
            table_html.append(u"<th>{0}</th>".format(_h(h)))
        table_html.append(u"</tr></thead><tbody>")
        for r in rows:
            table_html.append(u"<tr>{}</tr>".format(u"".join(u"<td>{}</td>".format(c) for c in r)))
        table_html.append(u"</tbody></table>")
        out.print_html(u"".join(table_html))
    else:
        out.print_html(u"<p>Нет обработанных стен для отображения.</p>")


if __name__ == "__main__":
    main()
