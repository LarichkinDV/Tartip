# -*- coding: utf-8 -*-
"""Определение ГЭСН для стен по таблице соответствий."""
import os
import sys
from collections import OrderedDict

from pyrevit import revit, DB, forms, script

# Добавляем путь к общим модулям расширения
THIS_DIR = os.path.dirname(__file__)
BASE_DIR = os.path.dirname(THIS_DIR)
LIB_DIR = os.path.join(BASE_DIR, "lib")
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from lib import config, gesn_rules, spec_keys_cache  # noqa: E402

# Имена используемых параметров
PARAM_REINFORCEMENT = u"Армирование"
PARAM_BRICK_SIZE = u"Размеры кирпича"
PARAM_STAGE = u"Стадия"
PARAM_STAGE_ALT = u"Стадия возведения"
STAGE_LABEL = (PARAM_STAGE or u"Стадия").lower()
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
    u"Объём": DB.BuiltInParameter.HOST_VOLUME_COMPUTED,
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


def _normalize_stage(value):
    """Приводит значение стадии к единому виду для сопоставления."""

    text = _t(value) or u""
    text = text.replace(u"\xa0", u" ")  # NBSP → space
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
    if not param:
        return 0.0, False
    return param.AsDouble() * FEET_TO_MM, True


def _get_thickness_mm(wall_type):
    param = wall_type.get_Parameter(PARAM_WIDTH)
    if not param:
        return 0.0, False
    return param.AsDouble() * FEET_TO_MM, True


def _get_type(wall):
    try:
        return wall.WallType
    except Exception:
        try:
            return revit.doc.GetElement(wall.GetTypeId())
        except Exception:
            return None


def _get_writable_param(element, names):
    """Возвращает первый доступный параметр из списка имен."""

    for name in names:
        param = element.LookupParameter(name)
        if param and not param.IsReadOnly:
            return param
    return None


def _resolve_type_info(wall):
    """Получает тип стены и гарантирует ненулевые имена семейства и типа."""

    wall_type = _get_type(wall)
    if wall_type is None:
        return None, u"(тип не определён)", u"(семейство не определено)"

    type_name = _t(getattr(wall_type, "Name", u""))
    if not type_name:
        param = wall_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if param:
            try:
                type_name = _t(param.AsString() or param.AsValueString())
            except Exception:
                type_name = u""
    if not type_name:
        type_name = u"(без имени)"

    family_name = _t(getattr(wall_type, "FamilyName", u""))
    if not family_name:
        try:
            family_name = _t(getattr(getattr(wall_type, "Family", None), "Name", u""))
        except Exception:
            family_name = u""
    if not family_name:
        family_name = u"(без семейства)"

    return wall_type, family_name, type_name


def _value_matches_conditions(value, conditions):
    """Проверяет значение по списку (оператор, число)."""

    if value is None:
        return False

    for op, limit in conditions:
        if op == ">" and not (value > limit):
            return False
        if op == ">=" and not (value >= limit):
            return False
        if op == "<" and not (value < limit):
            return False
        if op == "<=" and not (value <= limit):
            return False
        if op == "=" and not (abs(value - limit) <= 1e-6):
            return False
    return True


def _height_matches(rule, height_mm):
    if rule.height_conditions:
        return _value_matches_conditions(height_mm, rule.height_conditions)
    h_min = getattr(rule, "height_min_mm", None)
    h_max = getattr(rule, "height_max_mm", None)
    if h_min is None and h_max is None:
        return True
    if h_min is None:
        return height_mm <= h_max
    if h_max is None:
        return height_mm >= h_min
    return h_min <= height_mm <= h_max


def _get_extra_param_text(wall, wall_type, param_name):
    """Возвращает нормализованный текст значения доп. параметра (ФСБЦ и т.п.)."""

    value = _get_parameter_value(wall, param_name)
    if value is None and wall_type is not None:
        value = _get_parameter_value(wall_type, param_name)
    text = (_t(value) or u"").strip().lower()
    return text.replace(u"\xa0", u" ")


def _match_rules(
    rules,
    wall,
    wall_type,
    family_name,
    type_name,
    thickness_mm,
    height_mm,
    stage_text,
    reinforcement_text,
    brick_size,
):
    matched = []
    for rule in rules:
        if rule.family and rule.family != family_name:
            continue
        if rule.type_name and rule.type_name != type_name:
            continue
        has_rule_thickness = rule.thickness_mm is not None
        if has_rule_thickness:
            try:
                if abs(rule.thickness_mm - thickness_mm) > config.THICKNESS_TOLERANCE_MM:
                    continue
            except Exception:
                continue
        if not _height_matches(rule, height_mm):
            continue
        if rule.stage and rule.stage != stage_text:
            continue
        if rule.reinforcement and rule.reinforcement != reinforcement_text:
            continue
        if rule.brick_size and rule.brick_size != brick_size:
            continue
        extra_filters = getattr(rule, "extra_filters", None) or {}
        if extra_filters:
            extra_ok = True
            for name, expected in extra_filters.items():
                if not expected:
                    continue
                actual = _get_extra_param_text(wall, wall_type, name)
                if not actual or actual != expected:
                    extra_ok = False
                    break
            if not extra_ok:
                continue
        matched.append(rule)
    return matched


def _explain_no_match(
    rules,
    family_name,
    type_name,
    thickness_mm,
    height_mm,
    stage_text,
    reinforcement_text,
    brick_size,
    thickness_found=True,
    height_found=True,
    stage_found=True,
    reinf_found=True,
    brick_found=True,
):
    """Формирует человекочитаемую причину отсутствия совпадений."""

    stage_rules = list(rules)
    if not stage_rules:
        return u"Нет записей в БД (в таблице нет строк с кодами ГЭСН)"

    reasons = []

    family_rules = [r for r in stage_rules if not r.family or r.family == family_name]
    if not family_rules:
        reasons.append(u"семейство: {0}".format(family_name or u"(пусто)"))
    else:
        stage_rules = family_rules

    type_rules = [r for r in stage_rules if not r.type_name or r.type_name == type_name]
    if not type_rules:
        reasons.append(u"тип: {0}".format(type_name or u"(пусто)"))
    else:
        stage_rules = type_rules

    if not thickness_found:
        reasons.append(u"толщина: параметр не найден")
    else:
        thickness_rules = []
        for r in stage_rules:
            if r.thickness_mm is None:
                thickness_rules.append(r)
                continue
            try:
                if abs(r.thickness_mm - thickness_mm) <= config.THICKNESS_TOLERANCE_MM:
                    thickness_rules.append(r)
            except Exception:
                continue
        if not thickness_rules:
            reasons.append(u"толщина: {0:.1f} мм".format(thickness_mm))
        else:
            stage_rules = thickness_rules

    if not height_found:
        reasons.append(u"высота: параметр не найден")
    else:
        height_rules = [r for r in stage_rules if _height_matches(r, height_mm)]
        if not height_rules:
            expected_labels = []
            seen_expected = set()
            for r in stage_rules:
                label = (getattr(r, "height_label", u"") or u"").strip()
                if not label:
                    h_min = getattr(r, "height_min_mm", None)
                    h_max = getattr(r, "height_max_mm", None)
                    if h_min is None and h_max is None:
                        continue
                    if h_min is None:
                        label = u"<= {0:.1f} мм".format(h_max)
                    elif h_max is None:
                        label = u">= {0:.1f} мм".format(h_min)
                    else:
                        if abs(h_min) < 1e-6 and abs(h_max) < 1e-6:
                            continue
                        label = u"{0:.1f}-{1:.1f} мм".format(h_min, h_max)
                if label and label not in seen_expected:
                    seen_expected.add(label)
                    expected_labels.append(label)
            expected = u", ".join(expected_labels)
            reasons.append(
                u"высота {0:.1f} мм не соответствует ({1})".format(
                    height_mm,
                    expected or u"ожидание не задано",
                )
            )
        else:
            stage_rules = height_rules

    raw_stage = (stage_text or u"").strip()
    norm_stage = raw_stage.lower()
    stage_filtered = [r for r in stage_rules if not r.stage or r.stage == norm_stage]
    if not stage_filtered:
        display_stage = raw_stage or (
            u"параметр не найден" if not stage_found else u"(пусто)"
        )
        reasons.append(u"{0}: {1}".format(STAGE_LABEL, display_stage))
    else:
        stage_rules = stage_filtered

    norm_reinf = reinforcement_text or u""
    reinf_rules = [r for r in stage_rules if not r.reinforcement or r.reinforcement == norm_reinf]
    if not reinf_rules:
        msg = u"армирование: {0}".format(norm_reinf or (u"параметр не найден" if not reinf_found else u"(пусто)"))
        reasons.append(msg)
    else:
        stage_rules = reinf_rules

    norm_brick = (brick_size or u"").strip().lower()
    brick_rules = [r for r in stage_rules if not r.brick_size or r.brick_size == norm_brick]
    if not brick_rules:
        display_brick = brick_size or norm_brick
        msg = u"размеры кирпича: {0}".format(
            display_brick or (u"параметр не найден" if not brick_found else u"(пусто)")
        )
        reasons.append(msg)

    if not reasons:
        return u"Нет подходящей записи в БД"

    return u"Нет записей в БД ({0})".format(u"; ".join(reasons))


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


def _format_numeric_value(value, precision=5):
    """Форматирует число с округлением и заменой точки на запятую."""

    if value is None:
        return u""

    rounded = round(value, precision)
    template = u"{0:.{1}f}".format(rounded, precision)
    trimmed = template.rstrip("0").rstrip(".")
    if not trimmed:
        trimmed = u"0"
    return trimmed.replace(u".", u",")


def _calc_code_fragment(rule, volume_value):
    multiplier = rule.multiplier or 1.0
    normalized = volume_value / multiplier if multiplier else volume_value
    normalized_text = _format_numeric_value(normalized)
    return u"{0}[{1}]".format(rule.gesn_code, normalized_text)


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
    if getattr(rule, "volume_label", u""):
        parts.append(u"условие объёма: {0}".format(rule.volume_label))
    return u"; ".join(parts)


def _format_input_details(
    family_name,
    type_name,
    thickness_mm,
    thickness_found,
    height_mm,
    height_found,
    stage_text,
    stage_found,
    reinforcement_text,
    reinf_found,
    brick_size,
    brick_found,
):
    """Формирует строку со статусами исходных параметров."""

    items = []
    items.append(u"семейство={0}".format(family_name or u"(нет)"))
    items.append(u"тип={0}".format(type_name or u"(нет)"))

    if thickness_found:
        items.append(u"толщина={0:.1f} мм".format(thickness_mm))
    else:
        items.append(u"толщина=параметр не найден")

    if height_found:
        items.append(u"высота={0:.1f} мм".format(height_mm))
    else:
        items.append(u"высота=параметр не найден")

    if stage_found:
        items.append(u"{0}={1}".format(STAGE_LABEL, (stage_text or u"").strip() or u"(пусто)"))
    else:
        items.append(u"{0}=параметр не найден".format(STAGE_LABEL))

    if reinf_found:
        items.append(u"армирование={0}".format(reinforcement_text or u"(пусто)"))
    else:
        items.append(u"армирование=параметр не найден")

    if brick_found:
        items.append(u"размеры кирпича={0}".format(brick_size or u"(пусто)"))
    else:
        items.append(u"размеры кирпича=параметр не найден")

    return u"Данные: " + u"; ".join(items)


def _process_wall(wall, rules):
    """Обработка одной стены и запись результата/причины в параметр."""

    entry = {
        "id": getattr(getattr(wall, "Id", None), "IntegerValue", None),
        "cat": _t(getattr(getattr(wall, "Category", None), "Name", u"")),
        "type": None,
        "family": None,
        "message": None,
        "matched": False,
        "quantity_text": u"",
        "multiplier_text": u"",
        "gesn_text": u"",
    }

    target_param = _get_writable_param(wall, PARAM_GESN_OUTPUT)

    if not target_param:
        # Даже причину записать некуда
        entry["message"] = u"Нет доступного параметра для записи"
        return False, False, entry

    wall_type, family_name, type_name = _resolve_type_info(wall)
    entry["type"] = type_name
    entry["family"] = family_name

    if wall_type is None:
        reason = u"Не удалось определить тип стены"
        entry["message"] = reason
        return target_param.Set(reason), False, entry

    thickness_mm, thickness_found = _get_thickness_mm(wall_type)
    height_mm, height_found = _get_height_mm(wall)

    reinf_param = wall.LookupParameter(PARAM_REINFORCEMENT)
    reinforcement_text = _normalize_bool_text(_get_parameter_value(wall, PARAM_REINFORCEMENT)) if reinf_param else u""
    brick_param = wall.LookupParameter(PARAM_BRICK_SIZE)
    brick_size = _t(_get_parameter_value(wall, PARAM_BRICK_SIZE)) if brick_param else u""
    brick_size = (brick_size or u"").strip().lower()
    stage_text, stage_found = _get_stage_value(wall)

    input_details = _format_input_details(
        family_name,
        type_name,
        thickness_mm,
        thickness_found,
        height_mm,
        height_found,
        stage_text,
        stage_found,
        reinforcement_text,
        bool(reinf_param),
        brick_size,
        bool(brick_param),
    )

    if not thickness_found:
        reason = u"Не удалось определить толщину типа"
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    if not height_found:
        reason = u"Не удалось определить высоту стены"
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    matched_rules = _match_rules(
        rules,
        wall,
        wall_type,
        family_name,
        type_name,
        thickness_mm,
        height_mm,
        stage_text,
        reinforcement_text,
        brick_size,
    )
    if not matched_rules:
        reason = _explain_no_match(
            rules,
            entry["family"],
            entry["type"],
            thickness_mm,
            height_mm,
            stage_text,
            reinforcement_text,
            brick_size,
            thickness_found=thickness_found,
            height_found=height_found,
            stage_found=stage_found,
            reinf_found=bool(reinf_param),
            brick_found=bool(brick_param),
        )
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    def _rule_specificity(rule):
        score = 0
        if getattr(rule, "family", None):
            score += 1
        if getattr(rule, "type_name", None):
            score += 1
        if getattr(rule, "thickness_mm", None) is not None:
            score += 1
        if getattr(rule, "height_conditions", None):
            score += 1
        else:
            if getattr(rule, "height_min_mm", None) is not None or getattr(rule, "height_max_mm", None) is not None:
                score += 1
        if getattr(rule, "stage", None):
            score += 1
        if getattr(rule, "reinforcement", None):
            score += 1
        if getattr(rule, "brick_size", None):
            score += 1
        extra_filters = getattr(rule, "extra_filters", None) or {}
        if extra_filters:
            score += len(extra_filters)
        if getattr(rule, "volume_conditions", None):
            score += 1
        return score

    try:
        max_score = max(_rule_specificity(r) for r in matched_rules)
        matched_rules = [r for r in matched_rules if _rule_specificity(r) == max_score]
    except Exception:
        pass

    fragments_for_param = []
    fragments_for_report = []
    report_items = []
    last_volume_issue = None
    for rule in matched_rules:
        volume_value, unit_label = _get_volume_value(wall, rule)
        if volume_value is None:
            last_volume_issue = u"Не найден параметр объёма: {0}".format(rule.volume_param or u"?")
            continue
        if rule.volume_conditions and not _value_matches_conditions(volume_value, rule.volume_conditions):
            expected = rule.volume_label or u"условие объёма не задано"
            last_volume_issue = u"Объём {0:.3f} не попадает в диапазон ({1})".format(volume_value, expected)
            continue
        fragments_for_param.append(_calc_code_fragment(rule, volume_value))
        fragments_for_report.append(_format_rule_result(rule, volume_value, unit_label))
        report_items.append(
            {
                "gesn_code": rule.gesn_code,
                "volume_param": rule.volume_param or u"",
                "volume_value": volume_value,
                "unit_label": unit_label or u"",
                "multiplier": rule.multiplier or 1.0,
            }
        )

    if not fragments_for_param:
        reason = last_volume_issue or u"Не удалось вычислить объём"
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    unique_fragments = []
    seen_fragments = set()
    for fragment in fragments_for_param:
        if fragment in seen_fragments:
            continue
        seen_fragments.add(fragment)
        unique_fragments.append(fragment)

    unique_reports = []
    seen_reports = set()
    for fragment in fragments_for_report:
        if fragment in seen_reports:
            continue
        seen_reports.add(fragment)
        unique_reports.append(fragment)

    unique_items = []
    seen_item_keys = set()
    for item in report_items:
        key = (
            item.get("gesn_code"),
            item.get("volume_param"),
            round(item.get("volume_value") or 0.0, 6),
            item.get("unit_label"),
            item.get("multiplier"),
        )
        if key in seen_item_keys:
            continue
        seen_item_keys.add(key)
        unique_items.append(item)

    qty_parts = []
    seen_qty = set()
    for item in unique_items:
        qty_key = (
            item.get("volume_param"),
            round(item.get("volume_value") or 0.0, 6),
            item.get("unit_label"),
        )
        if qty_key in seen_qty:
            continue
        seen_qty.add(qty_key)
        vol_param = item.get("volume_param") or u""
        if vol_param:
            qty_parts.append(
                u"{0}: {1:.3f} {2}".format(
                    vol_param,
                    item.get("volume_value") or 0.0,
                    item.get("unit_label") or u"",
                )
            )
        else:
            qty_parts.append(
                u"{0:.3f} {1}".format(
                    item.get("volume_value") or 0.0,
                    item.get("unit_label") or u"",
                )
            )
    entry["quantity_text"] = u"; ".join(qty_parts)

    mult_parts = []
    seen_mult = set()
    for item in unique_items:
        mult = item.get("multiplier") or 1.0
        try:
            mult_text = int(mult) if float(mult).is_integer() else mult
        except Exception:
            mult_text = mult
        mult_str = _t(mult_text)
        if mult_str in seen_mult:
            continue
        seen_mult.add(mult_str)
        mult_parts.append(mult_str)
    entry["multiplier_text"] = u"; ".join(mult_parts)
    entry["gesn_text"] = u"; ".join(unique_fragments)

    entry["matched"] = True
    entry["message"] = u"{0} | {1}".format(u"; ".join(unique_reports), input_details)
    return target_param.Set(u"; ".join(unique_fragments)), True, entry


def _collect_walls():
    uidoc = revit.uidoc
    doc = revit.doc
    selection_ids = list(uidoc.Selection.GetElementIds()) if uidoc else []

    if selection_ids:
        elements = [doc.GetElement(eid) for eid in selection_ids]
        return [el for el in elements if getattr(el, "Category", None) and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Walls)]

    collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
    return list(collector)


def _ask_scope_choice():
    """Запрашивает у пользователя область обработки элементов."""

    options = OrderedDict()
    options[u"Выделенные элементы"] = "selection"
    options[u"Видимые элементы"] = "visible"
    options[u"Вся модель"] = "all"

    choice = forms.CommandSwitchWindow.show(
        options,
        message=u"Выберите область определения ГЭСН",
        width=500,
        height=250,
    )

    if not choice:
        return None

    return options.get(choice, choice)


def _is_wall(element):
    try:
        cat = getattr(element, "Category", None)
        return cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_Walls)
    except Exception:
        return False


def _collect_elements(scope):
    """Собирает стены по выбранной области."""

    uidoc = revit.uidoc
    doc = revit.doc

    if scope == "selection":
        selection_ids = list(uidoc.Selection.GetElementIds()) if uidoc else []
        if not selection_ids:
            forms.alert(u"Не выбраны элементы для обработки", exitscript=True)
            return []
        elements = [doc.GetElement(eid) for eid in selection_ids]
        return [el for el in elements if _is_wall(el)]

    if scope == "visible":
        view = getattr(doc, "ActiveView", None)
        if view is None:
            return []
        collector = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_Walls)
            .WhereElementIsNotElementType()
        )
        return list(collector)

    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
    )
    return list(collector)


def _select_source_and_update_cache():
    """Показывает диалог выбора источника правил (Excel/БД) и сохраняет выбор."""

    try:
        cache = spec_keys_cache.load_cache()
    except Exception:
        cache = None
    cache = cache or {}

    options = OrderedDict()
    options[u"Excel-файл с таблицей соответствия ГЭСН"] = "excel"
    options[u"База данных (SQL)"] = "db"

    choice = forms.CommandSwitchWindow.show(
        options,
        message=u"Выберите источник данных для правил ГЭСН",
        width=500,
        height=250,
    )

    if not choice:
        return None

    selected = options.get(choice, choice)

    if selected == "excel":
        initial_dir = None
        if cache.get("excel_path") and os.path.exists(cache["excel_path"]):
            initial_dir = os.path.dirname(cache["excel_path"])
        elif getattr(config, "EXCEL_PATH", None):
            try:
                initial_dir = os.path.dirname(config.EXCEL_PATH)
            except Exception:
                initial_dir = None

        excel_path = forms.pick_file(
            file_ext="xlsx",
            init_dir=initial_dir,
            title=u"Выбор Excel-файла с правилами ГЭСН",
        )

        if not excel_path:
            return None

        try:
            spec_keys_cache.save_cache(source_type="excel", excel_path=excel_path)
        except Exception:
            pass

        return "excel"

    if selected == "db":
        try:
            spec_keys_cache.save_cache(source_type="db")
        except Exception:
            pass
        return "db"

    return None


def _get_stage_value(wall):
    """Возвращает нормализованный текст стадии и флаг, что параметр найден."""

    for name in (PARAM_STAGE, PARAM_STAGE_ALT):
        stage_param = wall.LookupParameter(name)
        if stage_param:
            stage_value = _normalize_stage(_get_parameter_value(wall, name))
            return stage_value, True

    try:
        built_stage = wall.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
    except Exception:
        built_stage = None

    if built_stage:
        stage_value = None
        try:
            stage_value = built_stage.AsValueString()
        except Exception:
            stage_value = None
        if not stage_value:
            try:
                stage_value = built_stage.AsString()
            except Exception:
                stage_value = None
        return _normalize_stage(stage_value), True

    return u"", False


def _prepare_rules():
    """Загружает правила исходя из выбранного ранее источника."""

    cache = None
    try:
        cache = spec_keys_cache.load_cache()
    except Exception:
        cache = None

    if cache:
        source_type = cache.get("source_type")
        if source_type == "excel":
            excel_path = cache.get("excel_path")
            if excel_path:
                return gesn_rules.load_rules_from_excel(path=excel_path)
        elif source_type == "db":
            return gesn_rules.load_rules_from_db()

    return gesn_rules.load_rules_from_excel()


def main():
    out = script.get_output()

    scope_choice = _ask_scope_choice()
    if scope_choice is None:
        return

    source_choice = _select_source_and_update_cache()
    if source_choice is None:
        return
    try:
        rules = _prepare_rules()
    except Exception as exc:
        forms.alert(u"Не удалось загрузить таблицу правил: {0}".format(exc), exitscript=True)
        return

    elements = _collect_elements(scope_choice)
    if not elements:
        forms.alert(u"В модели не найдены стены для обработки", exitscript=True)
        return

    processed = 0
    updated = 0
    matched = 0
    entries = []

    with revit.Transaction(u"ТАРТИП: определить ГЭСН"):
        for wall in elements:
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
    summary_text = u"Обработано стен: {0}. Обновлено ГЭСН: {1}. Без подходящей записи: {2}.".format(
        processed,
        updated,
        not_matched,
    )

    # Выводим перечень элементов с найденными работами или причинами отсутствия
    try:
        out.clear()
    except Exception:
        pass

    out.print_html(u"<p><b>{0}</b></p>".format(_h(summary_text)))

    css = (
        u"<style>table.acbd{border-collapse:collapse;width:100%;margin:6px 0;color:#222;}"
        u"table.acbd th,table.acbd td{border:1px solid #d0d0d0;padding:4px 6px;}"
        u"table.acbd thead th{background:#e6e6e6;color:#101010;position:sticky;top:0;}</style>"
    )

    def _render_group(title, headers, rows, empty_message):
        out.print_html(u"<h3>{0}</h3>".format(_h(title)))
        if not rows:
            out.print_html(u"<p>{0}</p>".format(_h(empty_message)))
            return
        table_html = [u"<table class='acbd'>", u"<thead><tr>"]
        for h in headers:
            table_html.append(u"<th>{0}</th>".format(_h(h)))
        table_html.append(u"</tr></thead><tbody>")
        for r in rows:
            table_html.append(
                u"<tr>{0}</tr>".format(
                    u"".join(u"<td>{0}</td>".format(c) for c in r)
                )
            )
        table_html.append(u"</tbody></table>")
        out.print_html(u"".join(table_html))

    entries_with = [e for e in entries if e.get("matched")]
    entries_without = [e for e in entries if not e.get("matched")]

    rows_with = []
    for entry in entries_with:
        link = (
            out.linkify(DB.ElementId(entry["id"]), u"{}".format(entry["id"]))
            if entry.get("id") is not None
            else u""
        )
        rows_with.append(
            [
                link,
                _h(entry.get("cat") or u""),
                _h(entry.get("family") or u""),
                _h(entry.get("type") or u""),
                _h(entry.get("quantity_text") or u""),
                _h(entry.get("multiplier_text") or u""),
                _h(entry.get("gesn_text") or u""),
            ]
        )

    rows_without = []
    for entry in entries_without:
        link = (
            out.linkify(DB.ElementId(entry["id"]), u"{}".format(entry["id"]))
            if entry.get("id") is not None
            else u""
        )
        rows_without.append(
            [
                link,
                _h(entry.get("cat") or u""),
                _h(entry.get("family") or u""),
                _h(entry.get("type") or u""),
                _h(entry.get("message") or u""),
            ]
        )

    out.print_html(css)
    _render_group(
        title=u"ГЭСН определён",
        headers=[
            u"ID",
            u"Категория",
            u"Семейство",
            u"Тип",
            u"Кол-во/Объём",
            u"Кратность ед.изм. ГЭСН",
            u"Шифр ГЭСН",
        ],
        rows=rows_with,
        empty_message=u"Нет элементов с определённым шифром ГЭСН.",
    )
    _render_group(
        title=u"ГЭСН не определён",
        headers=[
            u"ID",
            u"Категория",
            u"Семейство",
            u"Тип",
            u"Результат",
        ],
        rows=rows_without,
        empty_message=u"Все элементы получили шифр ГЭСН.",
    )


if __name__ == "__main__":
    main()
