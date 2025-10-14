# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc = revit.doc
out = script.get_output()

# ---- Имена целевых параметров ----
P_COST_N_RATE = u"ACBD_Н_ЦенаЗаЕдИзм"
P_COST_F_RATE = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_LAB_N_RATE  = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_LAB_F_RATE  = u"ACBD_Ф_ТрудозатратыНаЕдИзм"

P_COST_N = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N  = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F  = u"ACBD_Ф_ТрудозатратыЭлемента"

# Возможные имена параметра ЕИ
UNIT_PARAM_NAMES = (
    u"ACBD_Н_ЕдиницаИзмерения",
    u"ACBD_ЕдиницаИзмерения",
    u"ADSK_Единица измерения",
    u"Единица измерения",
    u"Ед. изм.", u"Ед. изм",
    u"Unit",
)

# ------------------- утилиты текста/чисел -------------------
def _to_text(v):
    if v is None: return None
    try:
        if isinstance(v, unicode): return v
    except NameError:
        pass
    try: return unicode(v)
    except:
        try: return unicode(v.ToString())
        except:
            try: return unicode(str(v))
            except: return None

def _num(v):
    if v is None: return None
    s = _to_text(v)
    if not s: return None
    s = re.sub(u"[^0-9,.-]", u"", s.strip()).replace(u",", u".")
    try: return float(s)
    except: return None

# латинизация похожих кир. букв (Н→H, Ф→F, С→C и т.п.)
_cyr2lat = {u'Н':u'H',u'н':u'h',u'Ф':u'F',u'ф':u'f',u'С':u'C',u'с':u'c',u'А':u'A',u'а':u'a',u'В':u'B',u'в':u'b',
            u'Е':u'E',u'е':u'e',u'К':u'K',u'к':u'k',u'М':u'M',u'м':u'm',u'О':u'O',u'о':u'o',u'Р':u'P',u'р':u'p',
            u'Т':u'T',u'т':u't',u'Х':u'X',u'х':u'x',u'У':u'Y',u'у':u'y'}
def _latinize(s):
    s = _to_text(s) or u""
    return u"".join(_cyr2lat.get(ch, ch) for ch in s)

def _base_norm(name):
    s = _latinize(name).lower().replace(u"\u00a0", u"")
    for ch in (u" ", u"_", u".", u"-"): s = s.replace(ch, u"")
    s = s.replace(u"единицаизмерения", u"едизм")
    return s

# ------------------- доступ к параметрам (+fuzzy) -------------------
def _get_param(holder, name):
    if not holder: return None
    try:
        p = holder.LookupParameter(name)
        if p: return p
    except: pass
    target = _base_norm(name)
    pars = getattr(holder, "Parameters", None)
    if not pars: return None
    for p in pars:
        try:
            if _base_norm(p.Definition.Name) == target:
                return p
        except: pass
    return None

def _get_str(holder, name):
    p = _get_param(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.String:
            return _to_text(p.AsString())
        return _to_text(p.AsValueString())
    except: return None

def _get_double(holder, name):
    p = _get_param(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.Double: return p.AsDouble()
        if p.StorageType == DB.StorageType.String: return _num(p.AsString())
        return _num(p.AsValueString())
    except: return None

def _get_type(el):
    try: return doc.GetElement(el.GetTypeId())
    except: return None

def _get_rate(el, name):
    v = _get_str(el, name)
    if v is None:
        t = _get_type(el)
        if t: v = _get_str(t, name)
    return _num(v)

# ------------------- строительные элементы (вся модель) -------------------
_ALLOWED_CATS = List[DB.BuiltInCategory]([
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_Columns,              # архитектурные колонны
    DB.BuiltInCategory.OST_StructuralFraming,    # балки
    DB.BuiltInCategory.OST_StructuralFoundation,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Stairs,
    DB.BuiltInCategory.OST_Railings,
    DB.BuiltInCategory.OST_CurtainWallPanels,
    DB.BuiltInCategory.OST_CurtainWallMullions,
    DB.BuiltInCategory.OST_GenericModel,         # часто как строительные
])

def _building_elements_all():
    f_cats = DB.ElementMulticategoryFilter(_ALLOWED_CATS)
    col = (DB.FilteredElementCollector(doc)
             .WhereElementIsNotElementType()
             .WherePasses(f_cats))
    return [el for el in col if not getattr(el, "ViewSpecific", False)]

# ------------------- количество -------------------
def _bip(name):
    try: return getattr(DB.BuiltInParameter, name)
    except: return None

def _qty_area(el):
    for bip_name in ("HOST_AREA_COMPUTED", "ROOM_AREA"):
        bip = _bip(bip_name)
        if not bip: continue
        try:
            p = el.get_Parameter(bip)
            if p and p.AsDouble() and p.AsDouble() > 0:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.SquareMeters)
        except: pass
    for n in (u"Площадь", u"Area"):
        v = _get_double(el, n)
        if v is not None and v > 0: return v
    return 0.0

def _qty_volume(el):
    for bip_name in ("HOST_VOLUME_COMPUTED", "ROOM_VOLUME"):
        bip = _bip(bip_name)
        if not bip: continue
        try:
            p = el.get_Parameter(bip)
            if p and p.AsDouble() and p.AsDouble() > 0:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.CubicMeters)
        except: pass
    for n in (u"Объем", u"Объём", u"Volume"):
        v = _get_double(el, n)
        if v is not None and v > 0: return v
    return 0.0

def _qty_length(el):
    bip = _bip("CURVE_ELEM_LENGTH")
    if bip:
        try:
            p = el.get_Parameter(bip)
            if p and p.AsDouble() and p.AsDouble() > 0:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DB.UnitTypeId.Meters)
        except: pass
    for n in (u"Длина", u"Length"):
        v = _get_double(el, n)
        if v is not None and v > 0: return v
    return 0.0

def _qty_by_unit(el, unit_text):
    if not unit_text: return None
    ukey = (_latinize(unit_text) or u"").lower().replace(u" ", u"")
    if ukey in (u"м2", u"м²", u"m2"): return _qty_area(el)
    if ukey in (u"м3", u"м³", u"m3"): return _qty_volume(el)
    if ukey in (u"м", u"м.п.", u"мп", u"m"): return _qty_length(el)
    if ukey in (u"шт", u"шт.", u"штука", u"pcs"): return 1.0
    return None

def _auto_unit_and_qty(el):
    a = _qty_area(el)
    if a and a > 0: return u"м2", a, u"area"
    v = _qty_volume(el)
    if v and v > 0: return u"м3", v, u"volume"
    l = _qty_length(el)
    if l and l > 0: return u"м", l, u"length"
    return u"шт", 1.0, u"count"

def _get_unit_and_qty(el):
    for holder in (el, _get_type(el)):
        if not holder: continue
        for nm in UNIT_PARAM_NAMES:
            s = _get_str(holder, nm)
            if s:
                q = _qty_by_unit(el, s)
                if q is not None:
                    return s, q, u"param"
    return _auto_unit_and_qty(el)

# ------------------- запись (учёт «валюты») -------------------
def _is_currency_param(p):
    try:
        dt = p.Definition.GetDataType()
        if dt and dt.Equals(DB.SpecTypeId.Currency): return True
    except: pass
    try:
        return getattr(p.Definition, "ParameterType", None) == DB.ParameterType.Currency
    except: return False

def _fmt_currency(val):
    try:
        return DB.UnitFormatUtils.Format(doc, DB.SpecTypeId.Currency, float(val), False, False)
    except:
        return (u"{:,.2f}".format(float(val))).replace(u",", u" ").replace(u".", u",")

def _try_set_with_formats(p, value, is_currency):
    try:
        if p.Set(float(value)): return True
    except: pass
    if is_currency:
        for txt in (_fmt_currency(value),
                    u"{:.2f}".format(float(value)).replace(u".", u","),
                    u"{:,.2f}".format(float(value)).replace(u",", u" ").replace(u".", u","),
                    u"{:.2f}".format(float(value))):
            try:
                if p.SetValueString(txt): return True
            except: pass
    else:
        for txt in (u"{:.3f}".format(float(value)).replace(u".", u","), u"{:.3f}".format(float(value))):
            try:
                if p.SetValueString(txt): return True
            except: pass
    return False

def _set_number_any(el, target_name, value):
    # экземпляр
    p = _get_param(el, target_name)
    if p and not getattr(p, "IsReadOnly", False):
        if _try_set_with_formats(p, value, _is_currency_param(p)):
            return True, "instance"
    # тип
    t = _get_type(el)
    if t:
        pt = _get_param(t, target_name)
        if pt and not getattr(pt, "IsReadOnly", False):
            if _try_set_with_formats(pt, value, _is_currency_param(pt)):
                return True, "type"
    return False, "missing"

# ------------------- счётчики -------------------
w_inst = 0
w_type = 0
auto_units = 0
cost_written = 0
labor_written = 0

def _apply_value(el, target_name, value, is_cost):
    """Запись значения + обновление счётчиков (без вложенных функций)."""
    global w_inst, w_type, cost_written, labor_written
    ok, where = _set_number_any(el, target_name, value)
    if ok:
        if where == "instance": w_inst += 1
        else: w_type += 1
        if is_cost: cost_written += 1
        else: labor_written += 1
    return ok

# ------------------- расчёт по элементу -------------------
def _calc_and_set(el):
    global auto_units

    unit_txt, qty, src = _get_unit_and_qty(el)
    if src != u"param": auto_units += 1

    r_cn = _get_rate(el, P_COST_N_RATE)
    r_cf = _get_rate(el, P_COST_F_RATE)
    r_ln = _get_rate(el, P_LAB_N_RATE)
    r_lf = _get_rate(el, P_LAB_F_RATE)

    if r_cn is not None: _apply_value(el, P_COST_N, (r_cn or 0.0) * (qty or 0.0), True)
    if r_cf is not None: _apply_value(el, P_COST_F, (r_cf or 0.0) * (qty or 0.0), True)
    if r_ln is not None: _apply_value(el, P_LAB_N, (r_ln or 0.0) * (qty or 0.0), False)
    if r_lf is not None: _apply_value(el, P_LAB_F, (r_lf or 0.0) * (qty or 0.0), False)

# ------------------- основной цикл -------------------
elements = _building_elements_all()
if not forms.alert(u"Рассчитать для строительных элементов всего проекта?\nНайдено: {}".format(len(elements)), yes=True, no=True):
    script.exit()

with revit.Transaction(u"ACBD: расчёт (строительные элементы, весь проект)"):
    for el in elements:
        _calc_and_set(el)

forms.alert(
    u"Готово.\nЭлементов: {0}\nЗаписано в экземпляры: {1}\nЗаписано в типы: {2}\n"
    u"Стоимость записана: {3}\nТрудозатраты записаны: {4}\n"
    u"Авто-определений ЕИ: {5}".format(
        len(elements), w_inst, w_type, cost_written, labor_written, auto_units
    )
)
