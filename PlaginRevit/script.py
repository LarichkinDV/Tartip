# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, forms, script

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover - Python 3
    unicode = str  # type: ignore[assignment]

try:
    basestring  # type: ignore[name-defined]
except NameError:  # pragma: no cover - Python 3
    basestring = (str, bytes)  # type: ignore[assignment]

try:
    import System  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - pythonnet not available
    System = None

doc = revit.doc
out = script.get_output()

P_UNIT           = u"ACBD_Н_ЕдиницаИзмерения"
P_COST_N_RATE    = u"ACBD_Н_ЦенаЗаЕдИзм"
P_COST_F_RATE    = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_LAB_N_RATE     = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_LAB_F_RATE     = u"ACBD_Ф_ТрудозатратыНаЕдИзм"

P_COST_N         = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F         = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N          = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F          = u"ACBD_Ф_ТрудозатратыЭлемента"

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
    dotnet_string = getattr(System, "String", None) if System else None
    if dotnet_string and isinstance(value, dotnet_string):
        try:
            return unicode(value)
        except Exception:
            pass
    if hasattr(value, "ToString"):
        try:
            text = value.ToString()
        except Exception:
            text = None
        if text not in (None, ""):
            try:
                return unicode(text)
            except Exception:
                try:
                    return str(text)
                except Exception:
                    pass
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return None


def _num(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, bytes):
        text = _to_text(x)
    elif isinstance(x, basestring):  # type: ignore[arg-type]
        text = unicode(x)
    else:
        text = _to_text(x)
    if text is None:
        return None
    s = text.strip()
    if not s:
        return None
    s = re.sub(u"[^0-9,.-]", u"", s).replace(u",", u".")
    try:
        return float(s)
    except Exception:
        return None

def _get_param(el, name):
    try:
        return el.LookupParameter(name)
    except Exception:
        return None

def _get_str(el, name):
    p = _get_param(el, name)
    if not p:
        return None
    try:
        if p.StorageType == DB.StorageType.String:
            return _to_text(p.AsString())
        return _to_text(p.AsValueString())
    except Exception:
        return None

def _get_double(el, name):
    p = _get_param(el, name)
    if not p:
        return None
    try:
        if p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
        if p.StorageType == DB.StorageType.String:
            return _num(p.AsString())
        return _num(p.AsValueString())
    except Exception:
        return None

def _set_number(el, name, value):
    p = _get_param(el, name)
    if not p:
        return False
    if value is None:
        return False
    try:
        if p.StorageType == DB.StorageType.Double:
            return p.Set(float(value))
        elif p.StorageType == DB.StorageType.String:
            txt = u"{:,.2f}".format(float(value)).replace(u",", u" ")
            return p.Set(txt)
        else:
            return False
    except Exception:
        return False

def _qty_by_unit(el, unit_text):
    if not unit_text:
        return None
    u = unit_text.strip().lower().replace(u" ", u"")
    if u in (u"м2", u"м²"):
        for bip in (DB.BuiltInParameter.HOST_AREA_COMPUTED,
                    DB.BuiltInParameter.ROOM_AREA,
                    DB.BuiltInParameter.FLOOR_ATTR_AREA_COMPUTED):
            try:
                p = el.get_Parameter(bip)
                if p and p.AsDouble() and p.AsDouble() > 0:
                    return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.SquareMeters)
            except Exception:
                pass
        for n in (u"Площадь", u"Area"):
            v = _get_double(el, n)
            if v is not None and v > 0:
                return v
        return 0.0
    if u in (u"м3", u"м³"):
        for bip in (DB.BuiltInParameter.HOST_VOLUME_COMPUTED,
                    DB.BuiltInParameter.ROOM_VOLUME):
            try:
                p = el.get_Parameter(bip)
                if p and p.AsDouble() and p.AsDouble() > 0:
                    return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.CubicMeters)
            except Exception:
                pass
        for n in (u"Объем", u"Объём", u"Volume"):
            v = _get_double(el, n)
            if v is not None and v > 0:
                return v
        return 0.0
    if u in (u"м", u"м.п.", u"мп"):
        try:
            p = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
            if p and p.AsDouble() and p.AsDouble() > 0:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.Meters)
        except Exception:
            pass
        for n in (u"Длина", u"Length"):
            v = _get_double(el, n)
            if v is not None and v > 0:
                return v
        return 0.0
    if u in (u"шт", u"штука", u"шт.", u"pcs"):
        return 1.0
    return 1.0

def _calc_and_set(el, unit_text, rate_cost_n, rate_cost_f, rate_lab_n, rate_lab_f):
    qty = _qty_by_unit(el, unit_text)
    if qty is None:
        return False, u"Нет ед. изм."
    r_cn = _num(rate_cost_n)
    r_cf = _num(rate_cost_f)
    r_ln = _num(rate_lab_n)
    r_lf = _num(rate_lab_f)

    ok = False
    missing = []
    if r_cn is not None:
        ok |= _set_number(el, P_COST_N, r_cn * qty)
    else:
        missing.append(P_COST_N_RATE)
    if r_cf is not None:
        ok |= _set_number(el, P_COST_F, r_cf * qty)
    else:
        missing.append(P_COST_F_RATE)
    if r_ln is not None:
        ok |= _set_number(el, P_LAB_N, r_ln * qty)
    else:
        missing.append(P_LAB_N_RATE)
    if r_lf is not None:
        ok |= _set_number(el, P_LAB_F, r_lf * qty)
    else:
        missing.append(P_LAB_F_RATE)

    if not ok and missing:
        return False, u"Нет ставок: " + u", ".join(missing)
    return True, None

elems = [e for e in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
         if getattr(getattr(e, "Category", None), "CategoryType", None) == DB.CategoryType.Model]

updated = 0
skipped_no_unit = 0
skipped_other = 0

if not forms.alert(u"Пересчитать параметры стоимости и трудозатрат для {} элементов?\n"
                   u"Затрагиваемые параметры: \n • {}\n • {}\n • {}\n • {}"
                   .format(len(elems), P_COST_N, P_COST_F, P_LAB_N, P_LAB_F),
                   yes=True, no=True):
    script.exit()

with revit.Transaction(u"ACBD: Пересчёт стоимости/трудозатрат"):
    for el in elems:
        unit_text = _get_str(el, P_UNIT)
        if not unit_text:
            skipped_no_unit += 1
            continue
        cost_n = _get_str(el, P_COST_N_RATE)
        cost_f = _get_str(el, P_COST_F_RATE)
        lab_n  = _get_str(el, P_LAB_N_RATE)
        lab_f  = _get_str(el, P_LAB_F_RATE)

        ok, reason = _calc_and_set(el, unit_text, cost_n, cost_f, lab_n, lab_f)
        if ok:
            updated += 1
        else:
            if reason and reason.startswith(u"Нет ед. изм"):
                skipped_no_unit += 1
            else:
                skipped_other += 1

out.print_md(u"### Готово")
out.print_md(u"*Обновлено элементов:* **{0}**".format(updated))
out.print_md(u"*Пропущено (нет ед. изм.):* **{0}**".format(skipped_no_unit))
out.print_md(u"*Пропущено (нет ставок/прочее):* **{0}**".format(skipped_other))
forms.alert(u"Готово.\nОбновлено: {0}\nБез ед. изм.: {1}\nБез ставок/прочее: {2}"
            .format(updated, skipped_no_unit, skipped_other))
