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
PARAM_STAGE = u"Стадия"
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
    return rule.height_min_mm <= height_mm <= rule.height_max_mm


def _match_rules(
    rules,
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
        if abs(rule.thickness_mm - thickness_mm) > config.THICKNESS_TOLERANCE_MM:
            continue
        if not _height_matches(rule, height_mm):
            continue
        if rule.stage and rule.stage != stage_text:
            continue
        if rule.reinforcement and rule.reinforcement != reinforcement_text:
            continue
        if rule.brick_size and rule.brick_size != brick_size:
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
        thickness_rules = [r for r in stage_rules if abs(r.thickness_mm - thickness_mm) <= config.THICKNESS_TOLERANCE_MM]
        if not thickness_rules:
            reasons.append(u"толщина: {0:.1f} мм".format(thickness_mm))
        else:
            stage_rules = thickness_rules

    if not height_found:
        reasons.append(u"высота: параметр не найден")
    else:
        height_rules = [r for r in stage_rules if _height_matches(r, height_mm)]
        if not height_rules:
            expected = u", ".join(
                filter(
                    None,
                    [
                        getattr(r, "height_label", u"")
                        or u"{0:.1f}-{1:.1f} мм".format(
                            r.height_min_mm,
                            r.height_max_mm,
                        )
                        for r in stage_rules
                    ],
                )
            )
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
        reasons.append(u"стадия: {0}".format(display_stage))
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
        items.append(u"стадия={0}".format((stage_text or u"").strip() or u"(пусто)"))
    else:
        items.append(u"стадия=параметр не найден")

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
    brick_value_raw = _t(_get_parameter_value(wall, PARAM_BRICK_SIZE)) if brick_param else u""
    brick_size_normalized = (brick_value_raw or u"").strip().lower()
    stage_param = wall.LookupParameter(PARAM_STAGE)
    stage_value_raw = _t(_get_parameter_value(wall, PARAM_STAGE)) if stage_param else u""
    stage_text = (stage_value_raw or u"").strip().lower()

    input_details = _format_input_details(
        family_name,
        type_name,
        thickness_mm,
        thickness_found,
        height_mm,
        height_found,
        stage_value_raw,
        bool(stage_param),
        reinforcement_text,
        bool(reinf_param),
        brick_value_raw,
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
        family_name,
        type_name,
        thickness_mm,
        height_mm,
        stage_text,
        reinforcement_text,
        brick_size_normalized,
    )
    if not matched_rules:
        reason = _explain_no_match(
            rules,
            entry["family"],
            entry["type"],
            thickness_mm,
            height_mm,
            stage_value_raw,
            reinforcement_text,
            brick_value_raw,
            thickness_found=thickness_found,
            height_found=height_found,
            stage_found=bool(stage_param),
            reinf_found=bool(reinf_param),
            brick_found=bool(brick_param),
        )
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    fragments_for_param = []
    fragments_for_report = []
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

    if not fragments_for_param:
        reason = last_volume_issue or u"Не удалось вычислить объём"
        full_reason = u"{0} | {1}".format(reason, input_details)
        entry["message"] = full_reason
        return target_param.Set(full_reason), False, entry

    entry["matched"] = True
    entry["message"] = u"{0} | {1}".format(u"; ".join(fragments_for_report), input_details)
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
    excel_path = forms.pick_file(
        file_ext="xlsx",
        multi_file=False,
        title=u"Выберите файл Excel с таблицей соответствий ГЭСН",
    )
    if not excel_path:
        forms.alert(u"Файл с правилами не выбран. Операция отменена.", exitscript=True)
        return

    try:
        rules = load_rules_from_excel(path=excel_path)
    except Exception as exc:
        forms.alert(
            u"Не удалось загрузить таблицу правил из файла:\n{0}\n\nОшибка: {1}".format(
                excel_path,
                exc,
            ),
            exitscript=True,
        )
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
            u"<style>table.acbd{border-collapse:collapse;width:100%;margin:6px 0;color:#222;}"
            u"table.acbd th,table.acbd td{border:1px solid #d0d0d0;padding:4px 6px;}"
            u"table.acbd thead th{background:#e6e6e6;color:#101010;position:sticky;top:0;}</style>"
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
