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
PARAM_GESN_CODE = u"Шифр ГЭСН"

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
    return u"{0}[{1}/{2}]".format(rule.gesn_code, volume_param, int(multiplier) if multiplier.is_integer() else multiplier)


def _process_wall(wall, rules):
    wall_type = _get_type(wall)
    if wall_type is None:
        return False, False

    thickness_mm = _get_thickness_mm(wall_type)
    height_mm = _get_height_mm(wall)
    reinforcement_text = _normalize_bool_text(_get_parameter_value(wall, PARAM_REINFORCEMENT))
    brick_size = _t(_get_parameter_value(wall, PARAM_BRICK_SIZE))

    matched_rules = _match_rules(wall, rules, thickness_mm, height_mm, reinforcement_text, brick_size)
    if not matched_rules:
        return True, False

    fragments = []
    for rule in matched_rules:
        volume_value, unit_label = _get_volume_value(wall, rule)
        if volume_value is None:
            continue
        fragments.append(_calc_code_fragment(rule, volume_value))

    if not fragments:
        return True, False

    target_param = wall.LookupParameter(PARAM_GESN_CODE)
    if not target_param or target_param.IsReadOnly:
        return False, False

    return target_param.Set(u"; ".join(fragments)), True


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

    with revit.Transaction(u"ТАРТИП: определить ГЭСН"):
        for wall in walls:
            ok, has_match = _process_wall(wall, rules)
            if ok:
                processed += 1
                if has_match:
                    matched += 1
                    updated += 1
                elif config.CLEAR_CODE_WHEN_MISS:
                    param = wall.LookupParameter(PARAM_GESN_CODE)
                    if param and not param.IsReadOnly:
                        param.Set(u"")
            else:
                out.print_html(u"Не удалось обновить стену: {0}".format(_t(wall)))

    not_matched = processed - matched
    forms.alert(
        u"Обработано стен: {0}\nОбновлено ГЭСН: {1}\nБез подходящей записи: {2}".format(
            processed, updated, not_matched
        )
    )


if __name__ == "__main__":
    main()
