# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, forms, script


doc = revit.doc
out = script.get_output()

P_UNIT = u"ACBD_Н_ЕдиницаИзмерения"
P_COST_N_RATE = u"ACBD_Н_ЦенаЗаЕдИзм"
P_COST_F_RATE = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_LAB_N_RATE = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_LAB_F_RATE = u"ACBD_Ф_ТрудозатратыНаЕдИзм"

P_COST_N = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F = u"ACBD_Ф_ТрудозатратыЭлемента"


def _to_text(value):
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, bytes):
        for encoding in ("utf-8", "cp1251"):
            try:
                return value.decode(encoding)
            except UnicodeDecodeError:
                continue
        return value.decode("utf-8", "ignore")
    if hasattr(value, "ToString"):
        try:
            value = value.ToString()
        except Exception:
            value = None
    try:
        return str(value) if value is not None else None
    except Exception:
        return None


def _num(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _to_text(value)
    if not text:
        return None
    text = text.strip()
    if not text:
        return None
    text = re.sub(u"[^0-9,.-]", u"", text).replace(u",", u".")
    try:
        return float(text)
    except Exception:
        return None


def _get_param(el, name):
    try:
        return el.LookupParameter(name)
    except Exception:
        return None


def _get_str(el, name):
    param = _get_param(el, name)
    if not param:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            return _to_text(param.AsString())
        return _to_text(param.AsValueString())
    except Exception:
        return None


def _get_double(el, name):
    param = _get_param(el, name)
    if not param:
        return None
    try:
        if param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        if param.StorageType == DB.StorageType.String:
            return _num(param.AsString())
        return _num(param.AsValueString())
    except Exception:
        return None


def _set_number(el, name, value):
    param = _get_param(el, name)
    if not param or value is None:
        return False
    try:
        if param.StorageType == DB.StorageType.Double:
            return param.Set(float(value))
        if param.StorageType == DB.StorageType.String:
            text = u"{:,.2f}".format(float(value)).replace(u",", u" ")
            return param.Set(text)
    except Exception:
        return False
    return False


def _qty_by_unit(el, unit_text):
    if not unit_text:
        return None
    unit_key = unit_text.strip().lower().replace(u" ", u"")
    if unit_key in (u"м2", u"м²"):
        for bip in (
            DB.BuiltInParameter.HOST_AREA_COMPUTED,
            DB.BuiltInParameter.ROOM_AREA,
            DB.BuiltInParameter.FLOOR_ATTR_AREA_COMPUTED,
        ):
            try:
                param = el.get_Parameter(bip)
                if param and param.AsDouble() and param.AsDouble() > 0:
                    return DB.UnitUtils.ConvertFromInternalUnits(
                        param.AsDouble(), DB.UnitTypeId.SquareMeters
                    )
            except Exception:
                continue
        for name in (u"Площадь", u"Area"):
            value = _get_double(el, name)
            if value is not None and value > 0:
                return value
        return 0.0
    if unit_key in (u"м3", u"м³"):
        for bip in (
            DB.BuiltInParameter.HOST_VOLUME_COMPUTED,
            DB.BuiltInParameter.ROOM_VOLUME,
        ):
            try:
                param = el.get_Parameter(bip)
                if param and param.AsDouble() and param.AsDouble() > 0:
                    return DB.UnitUtils.ConvertFromInternalUnits(
                        param.AsDouble(), DB.UnitTypeId.CubicMeters
                    )
            except Exception:
                continue
        for name in (u"Объем", u"Объём", u"Volume"):
            value = _get_double(el, name)
            if value is not None and value > 0:
                return value
        return 0.0
    if unit_key in (u"м", u"м.п.", u"мп"):
        try:
            param = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
            if param and param.AsDouble() and param.AsDouble() > 0:
                return DB.UnitUtils.ConvertFromInternalUnits(
                    param.AsDouble(), DB.UnitTypeId.Meters
                )
        except Exception:
            pass
        for name in (u"Длина", u"Length"):
            value = _get_double(el, name)
            if value is not None and value > 0:
                return value
        return 0.0
    if unit_key in (u"шт", u"штука", u"шт.", u"pcs"):
        return 1.0
    return 1.0


def _calc_and_set(el, unit_text, rate_cost_n, rate_cost_f, rate_lab_n, rate_lab_f):
    quantity = _qty_by_unit(el, unit_text)
    if quantity is None:
        return False, u"Нет ед. изм."

    rate_cost_n = _num(rate_cost_n)
    rate_cost_f = _num(rate_cost_f)
    rate_lab_n = _num(rate_lab_n)
    rate_lab_f = _num(rate_lab_f)

    updated = False
    missing = []

    if rate_cost_n is not None:
        updated |= _set_number(el, P_COST_N, rate_cost_n * quantity)
    else:
        missing.append(P_COST_N_RATE)

    if rate_cost_f is not None:
        updated |= _set_number(el, P_COST_F, rate_cost_f * quantity)
    else:
        missing.append(P_COST_F_RATE)

    if rate_lab_n is not None:
        updated |= _set_number(el, P_LAB_N, rate_lab_n * quantity)
    else:
        missing.append(P_LAB_N_RATE)

    if rate_lab_f is not None:
        updated |= _set_number(el, P_LAB_F, rate_lab_f * quantity)
    else:
        missing.append(P_LAB_F_RATE)

    if not updated and missing:
        return False, u"Нет ставок: " + u", ".join(missing)
    return True, None


elements = [
    el
    for el in DB.FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
    if getattr(getattr(el, "Category", None), "CategoryType", None)
    == DB.CategoryType.Model
]

updated = 0
skipped_no_unit = 0
skipped_other = 0

if not forms.alert(
    u"Пересчитать параметры стоимости и трудозатрат для {} элементов?\n"
    u"Затрагиваемые параметры: \n • {}\n • {}\n • {}\n • {}".format(
        len(elements), P_COST_N, P_COST_F, P_LAB_N, P_LAB_F
    ),
    yes=True,
    no=True,
):
    script.exit()

with revit.Transaction(u"ACBD: Пересчёт стоимости/трудозатрат"):
    for element in elements:
        unit_text = _get_str(element, P_UNIT)
        if not unit_text:
            skipped_no_unit += 1
            continue

        cost_n = _get_str(element, P_COST_N_RATE)
        cost_f = _get_str(element, P_COST_F_RATE)
        lab_n = _get_str(element, P_LAB_N_RATE)
        lab_f = _get_str(element, P_LAB_F_RATE)

        ok, reason = _calc_and_set(element, unit_text, cost_n, cost_f, lab_n, lab_f)
        if ok:
            updated += 1
        elif reason and reason.startswith(u"Нет ед. изм"):
            skipped_no_unit += 1
        else:
            skipped_other += 1

out.print_md(u"### Готово")
out.print_md(u"*Обновлено элементов:* **{0}**".format(updated))
out.print_md(u"*Пропущено (нет ед. изм.):* **{0}**".format(skipped_no_unit))
out.print_md(u"*Пропущено (нет ставок/прочее):* **{0}**".format(skipped_other))

summary_message = u"Готово.\nОбновлено: {0}\nБез ед. изм.: {1}\nБез ставок/прочее: {2}".format(
    updated,
    skipped_no_unit,
    skipped_other,
)

forms.alert(
    summary_message,
)
