# -*- coding: utf-8 -*-
import re, os, datetime, zipfile
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List as CsList
from System.Windows import Window, WindowStyle, ResizeMode, Thickness, HorizontalAlignment
from System.Windows.Controls import StackPanel, TextBlock, RadioButton, CheckBox, Button, Orientation

doc = revit.doc
out = script.get_output()

# ---- ACBD параметры ----
# Источник (ТИП)
P_UNIT_T    = u"ACBD_ЕдиницаИзмерения"
P_RATE_CN_T = u"ACBD_Н_ЦенаЗаЕдИзм"
P_RATE_CF_T = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_RATE_LN_T = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_RATE_LF_T = u"ACBD_Ф_ТрудозатратыНаЕдИзм"
# Приёмники (ЭКЗЕМПЛЯР)
P_COST_N_I  = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F_I  = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N_I   = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F_I   = u"ACBD_Ф_ТрудозатратыЭлемента"

try:
    text_type = unicode
except NameError:
    text_type = str

_C2L = {u"А":u"A",u"а":u"a",u"В":u"B",u"в":u"b",u"С":u"C",u"с":u"c",u"Е":u"E",u"е":u"e",
        u"Н":u"H",u"н":u"h",u"К":u"K",u"к":u"k",u"М":u"M",u"м":u"m",u"О":u"O",u"о":u"o",
        u"Р":u"P",u"р":u"p",u"Т":u"T",u"т":u"t",u"Х":u"X",u"х":u"x",u"У":u"Y",u"у":u"y",u"Ф":u"F",u"ф":u"f"}

def _t(x):
    if x is None: return None
    try:
        if isinstance(x, text_type): return x
    except: pass
    try: return text_type(x)
    except:
        try: return text_type(x.ToString())
        except: return None

def _fold(s):
    s = (_t(s) or u"").replace(u"\u00a0", u" ").strip()
    return u"".join(_C2L.get(ch, ch) for ch in s).lower()

def _num(v):
    if v is None: return None
    s = _t(v)
    if not s: return None
    s = re.sub(u"[^0-9,.-]", u"", s.strip()).replace(u",", u".")
    try: return float(s)
    except: return None

def _fmt_money(v):
    try:
        return DB.UnitFormatUtils.Format(doc, DB.SpecTypeId.Currency, float(v), False, False)
    except:
        return (u"{:,.2f}".format(float(v))).replace(u",", u" ").replace(u".", u",")

def _fmt_num(v, nd=3):
    try:
        s = u"{:,.%df}" % nd
        return s.format(float(v)).replace(u",", u" ").replace(u".", u",")
    except:
        return _t(v) or u""

def _lp(holder, name):
    if not holder: return None
    try:
        p = holder.LookupParameter(name)
        if p: return p
    except: pass
    want = _fold(name)
    try:
        for p in holder.Parameters:
            if _fold(getattr(p.Definition, "Name", u"")) == want:
                return p
    except: pass
    return None

def _eltype(el):
    try: return doc.GetElement(el.GetTypeId())
    except: return None

def _type_name(el):
    et = _eltype(el)
    if et:
        n = _t(getattr(et, "Name", None))
        if n: return n
        try:
            p = et.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                s = _t(p.AsString())
                if s: return s
        except: pass
    try:
        p = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            vs = _t(p.AsValueString())
            if vs: return vs
            etid = p.AsElementId()
            if etid and etid.IntegerValue>0:
                et2 = doc.GetElement(etid)
                if et2:
                    n2 = _t(getattr(et2,"Name",None))
                    if n2: return n2
    except: pass
    try:
        sym = getattr(el, "Symbol", None)
        if sym:
            fam = getattr(sym, "Family", None)
            fname = _t(getattr(fam, "Name", None)) if fam else None
            sname = _t(getattr(sym, "Name", None))
            if fname or sname:
                if fname and sname and fname != sname:
                    return u"{} : {}".format(fname, sname)
                return fname or sname
    except: pass
    return u""

def _get_str_from(holder, name):
    p = _lp(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.String: return _t(p.AsString())
        return _t(p.AsValueString())
    except: return None

def _get_num_from(holder, name):
    p = _lp(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.Double: return p.AsDouble()
        if p.StorageType == DB.StorageType.String: return _num(p.AsString())
        return _num(p.AsValueString())
    except: return None

def _inst_param(el, name): return _lp(el, name)

def _is_currency(p):
    try:
        dt = p.Definition.GetDataType()
        if dt and dt.Equals(DB.SpecTypeId.Currency): return True
    except: pass
    try:
        return getattr(p.Definition,"ParameterType",None) == DB.ParameterType.Currency
    except: return False

def _try_set_number(p, value):
    try:
        if p.Set(float(value)): return True
    except: pass
    if _is_currency(p):
        variants = (
            _fmt_money(value),
            u"{:.2f}".format(float(value)).replace(u".", u","),
            u"{:,.2f}".format(float(value)).replace(u",", u" ").replace(u".", u","),
            u"{:.2f}".format(float(value)),
        )
    else:
        variants = (
            u"{:.3f}".format(float(value)).replace(u".", u","),
            u"{:.3f}".format(float(value)),
        )
    for s in variants:
        try:
            if p.SetValueString(s): return True
        except: pass
    return False

def _set_inst_number(el, name, value):
    p = _inst_param(el, name)
    if not p or getattr(p, "IsReadOnly", False): return False
    return _try_set_number(p, value)

# ---- сбор элементов ----
ALLOWED = CsList[DB.BuiltInCategory]([
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_Columns,
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralFoundation,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Stairs,
    DB.BuiltInCategory.OST_Railings,
    DB.BuiltInCategory.OST_CurtainWallPanels,
    DB.BuiltInCategory.OST_CurtainWallMullions,
    DB.BuiltInCategory.OST_GenericModel,
])

def _collect_all():
    f = DB.ElementMulticategoryFilter(ALLOWED)
    col = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(f)
    return [el for el in col
            if getattr(el, "ViewSpecific", False) is False
            and getattr(getattr(el,"Category",None),"CategoryType",None) == DB.CategoryType.Model]

def _collect_visible(view):
    f = DB.ElementMulticategoryFilter(ALLOWED)
    col = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().WherePasses(f)
    return [el for el in col
            if getattr(getattr(el,"Category",None),"CategoryType",None) == DB.CategoryType.Model]

# ---- количества из экземпляра ----
def _get_double_si(el, names, unit_tid):
    if not isinstance(names,(list,tuple)): names=(names,)
    for nm in names:
        p = _inst_param(el, nm)
        if not p: continue
        try:
            if p.StorageType == DB.StorageType.Double:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), unit_tid)
            if p.StorageType == DB.StorageType.String:
                v = _num(p.AsString())
                if v is not None: return v
            vs = _t(p.AsValueString())
            v = _num(vs)
            if v is not None: return v
        except: pass
    return None

def _qty(el, unit_text):
    if not unit_text: return None
    key = (_t(unit_text) or u"").lower().replace(u"\u00a0",u" ").replace(u" ",u"").strip()
    if key in (u"квм",u"кв.м",u"м2",u"м²",u"m2",u"sqm"): key = u"м2"
    if key in (u"кубм",u"куб.м",u"м3",u"м³",u"m3",u"cbm"): key = u"м3"
    if key in (u"м",u"мп",u"м.п",u"м.п.",u"п.м",u"pm",u"rm"): key = u"м"
    if key in (u"шт",u"шт.",u"штука",u"pcs",u"pc"): key = u"шт"
    if key == u"м2":
        v = _get_double_si(el, (u"Area",u"Площадь"), DB.UnitTypeId.SquareMeters); return 0.0 if v is None else v
    if key == u"м3":
        v = _get_double_si(el, (u"Volume",u"Объем",u"Объём"), DB.UnitTypeId.CubicMeters); return 0.0 if v is None else v
    if key == u"м":
        v = _get_double_si(el, (u"Length",u"Длина"), DB.UnitTypeId.Meters); return 0.0 if v is None else v
    if key == u"шт": return 1.0
    return None

# ---- стадии ----
ST_EXIST  = u"Существующие"
ST_DEMOL  = u"Демонтаж"
ST_NEW    = u"Новые конструкции"
ST_OTHER  = u"Прочее"

def _phase_names(el):
    cr = None; dm = None
    try:
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if p:
            pid = p.AsElementId()
            if pid and pid.IntegerValue>0:
                ph = doc.GetElement(pid); cr = _t(getattr(ph,"Name",None))
    except: pass
    try:
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
        if p:
            pid = p.AsElementId()
            if pid and pid.IntegerValue>0:
                ph = doc.GetElement(pid); dm = _t(getattr(ph,"Name",None))
    except: pass
    return cr, dm

def _nru(s):  # to lower ru
    return (_t(s) or u"").strip().lower().replace(u"\u00a0", u" ")

def _stage_bucket(el):
    cr, dm = _phase_names(el)
    crl, dml = _nru(cr), _nru(dm)
    if (dml == u"демонтаж") and (crl == u"существующие"): return ST_DEMOL
    if (not dml) and (crl == u"новая конструкция"):       return ST_NEW
    if crl == u"существующие":                             return ST_EXIST
    if u"демонтаж" in dml and u"существ" in crl:           return ST_DEMOL
    if (not dml) and (u"нов" in crl):                      return ST_NEW
    if u"существ" in crl:                                  return ST_EXIST
    return ST_OTHER

# ---- расчёт одного элемента ----
def _calc_element(el, buckets_calc, buckets_skip, totals):
    et     = _eltype(el)
    tname  = _type_name(el) or u"(без имени типа)"
    cat    = _t(getattr(getattr(el,"Category",None),"Name",u"(нет категории)"))
    eid    = getattr(getattr(el,"Id",None),"IntegerValue",None)
    stage  = _stage_bucket(el)

    unit_text = _get_str_from(et, P_UNIT_T)
    if not unit_text or not _t(unit_text).strip():
        buckets_skip.setdefault(stage, {}).setdefault(tname, []).append(
            dict(id=eid, cat=cat, tname=tname, reason=u"ЕИ пуста в типе")
        )
        return False

    q    = _qty(el, unit_text)
    if q is None:
        buckets_skip.setdefault(stage, {}).setdefault(tname, []).append(
            dict(id=eid, cat=cat, tname=tname, reason=u"ЕИ '{}' не распознана".format(unit_text))
        )
        return False

    r_cn = _get_num_from(et, P_RATE_CN_T)
    r_cf = _get_num_from(et, P_RATE_CF_T)
    r_ln = _get_num_from(et, P_RATE_LN_T)
    r_lf = _get_num_from(et, P_RATE_LF_T)

    ok_any = False
    cost_n = cost_f = lab_n = lab_f = None

    if r_cn is not None:
        cost_n = (r_cn or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_COST_N_I, cost_n)
        totals["N"] += float(cost_n)
    if r_cf is not None:
        cost_f = (r_cf or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_COST_F_I, cost_f)
        totals["F"] += float(cost_f)
    if r_ln is not None:
        lab_n = (r_ln or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_LAB_N_I, lab_n)
        totals["LN"] += float(lab_n)
    if r_lf is not None:
        lab_f = (r_lf or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_LAB_F_I, lab_f)
        totals["LF"] += float(lab_f)

    if not ok_any:
        buckets_skip.setdefault(stage, {}).setdefault(tname, []).append(
            dict(id=eid, cat=cat, tname=tname, reason=u"Нет ставок в типе")
        )
        return False

    # положим в рассчитанные
    bucket = buckets_calc.setdefault(stage, {}).setdefault(tname, dict(
        sumN=0.0, sumF=0.0, sumLN=0.0, sumLF=0.0, count=0, items=[]
    ))
    if cost_n is not None: bucket["sumN"]  += float(cost_n)
    if cost_f is not None: bucket["sumF"]  += float(cost_f)
    if lab_n  is not None: bucket["sumLN"] += float(lab_n)
    if lab_f  is not None: bucket["sumLF"] += float(lab_f)
    bucket["count"] += 1
    bucket["items"].append(dict(
        id=eid, cat=cat, tname=tname, unit=unit_text, qty=q,
        rcn=r_cn, rcf=r_cf, rln=r_ln, rlf=r_lf,
        cn=cost_n, cf=cost_f, ln=lab_n, lf=lab_f
    ))
    return True

# ---- HTML рендер (панель вывода) ----
def _h(s):
    if s is None: return u""
    s = _t(s)
    return (s.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
             .replace(u">", u"&gt;").replace(u'"', u"&quot;"))

def _table(headers, rows, align=None, safe=None):
    safe = set(safe or [])
    align = align or []
    th = []
    for i,h in enumerate(headers):
        a = align[i] if i<len(align) else "left"
        th.append(u'<th style="text-align:{}">{}</th>'.format(a, _h(h)))
    trs=[]
    for r in rows:
        tds=[]
        for c,val in enumerate(r):
            a = align[c] if c<len(align) else "left"
            txt = u"{}".format(val) if c in safe else _h(val)
            tds.append(u'<td style="text-align:{}">{}</td>'.format(a, txt))
        trs.append(u"<tr>{}</tr>".format(u"".join(tds)))
    return u"<table class='acbd'><thead><tr>{}</tr></thead><tbody>{}</tbody></table>".format(u"".join(th), u"".join(trs))

def _render_report(calc_map, skip_map, totals, processed, okcnt):
    try: out.clear()
    except: pass

    css = u"""
    <style>
      .acbd-wrap{font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1b1b1b;}
      .acbd h1{font-size:20px;margin:8px 0 8px;}
      .acbd h2{font-size:16px;margin:12px 0 8px;}
      .acbd h3{font-size:14px;margin:8px 0 6px;}
      .pill{display:inline-block;background:#eef3ff;border:1px solid #cdd9ff;color:#1f3b8f;padding:2px 6px;border-radius:10px;font-size:12px}
      table.acbd{border-collapse:collapse;width:100%;margin:6px 0 10px;}
      table.acbd th,table.acbd td{border:1px solid #d0d0d0;padding:6px 8px}
      table.acbd thead th{position:sticky;top:0;background:#f6f6f6;z-index:1}
      table.acbd tbody tr:nth-child(odd){background:#fafafa}
      details{margin:4px 0 8px 0;border:1px solid #e0e0e0;border-radius:6px;padding:6px 10px;background:#fff}
      details>summary{cursor:pointer;font-weight:600}
      .muted{color:#666}
      .mono{font-family:Consolas,Menlo,monospace}
    </style>"""

    html = [u'<div class="acbd-wrap">', css, u'<div class="acbd">']

    # ===== Заголовок: 4 ключевых суммы
    html.append(u"<h1>Стоимость проектируемого объекта</h1>")
    key_rows = [
        [u"Нормативная оценка стоимости (ГЭСН)",  _fmt_money(totals["N"])],
        [u"Опытная оценка стоимости",             _fmt_money(totals["F"])],
        [u"Нормативная оценка трудозатрат",       _fmt_num(totals["LN"])],
        [u"Опытная оценка трудозатрат",           _fmt_num(totals["LF"])],
    ]
    html.append(_table([u"Метрика", u"Значение"], key_rows, align=["left","right"]))

    html.append(u'<p class="muted">Обработано элементов: {} &nbsp;&nbsp; С расчётом: {} &nbsp;&nbsp; Пропущено: {}</p>'
                .format(processed, okcnt, processed - okcnt))

    # ===== Рассчитанные
    html.append(u"<h1>Итоги по стадиям и типам — рассчитанные</h1>")
    if not calc_map:
        html.append(u"<p><i>Нет рассчитанных элементов.</i></p>")
    else:
        for stage in (ST_EXIST, ST_DEMOL, ST_NEW, ST_OTHER):
            stage_types = calc_map.get(stage)
            if not stage_types: continue
            html.append(u'<details open><summary>{}</summary>'.format(_h(stage)))
            # типы — по имени А→Я
            for tname in sorted(stage_types.keys(), key=lambda s: _fold(s)):
                data = stage_types[tname]
                # заголовок по типу
                head = u"{}  —  x{}  |  Н: {}  |  Ф: {}".format(
                    _h(tname), data["count"], _fmt_money(data["sumN"]), _fmt_money(data["sumF"])
                )
                html.append(u'<details><summary>{}</summary>'.format(head))
                # элементы — по суммарной стоимости (Н+Ф) убыв.
                items = sorted(list(data["items"]),
                               key=lambda it: (float(it["cn"] or 0.0)+float(it["cf"] or 0.0)),
                               reverse=True)
                rows = []
                for it in items:
                    link = out.linkify(DB.ElementId(it["id"]), u"{}".format(it["id"]))
                    rows.append([
                        link, _h(it["cat"]), _h(it["unit"]),
                        _fmt_num(it["qty"]), _fmt_money(it["rcn"] or 0.0), _fmt_money(it["rcf"] or 0.0),
                        _fmt_money(it["cn"] or 0.0), _fmt_money(it["cf"] or 0.0),
                        _fmt_num(it["ln"] or 0.0), _fmt_num(it["lf"] or 0.0)
                    ])
                html.append(_table(
                    [u"ID", u"Категория", u"ЕИ", u"Кол-во", u"Н цена/ед", u"Ф цена/ед",
                     u"Н стоимость", u"Ф стоимость", u"Н труд.", u"Ф труд."],
                    rows, align=["right","left","left","right","right","right","right","right","right","right"], safe=[0]
                ))
                html.append(u'</details>')
            html.append(u'</details>')

    # ===== Нерассчитанные
    html.append(u"<h1>Итоги по стадиям и типам — нерассчитанные</h1>")
    if not skip_map:
        html.append(u"<p><i>Все элементы рассчитаны.</i></p>")
    else:
        for stage in (ST_EXIST, ST_DEMOL, ST_NEW, ST_OTHER):
            stage_types = skip_map.get(stage)
            if not stage_types: continue
            html.append(u'<details><summary>{}</summary>'.format(_h(stage)))
            for tname in sorted(stage_types.keys(), key=lambda s: _fold(s)):
                items = stage_types[tname]
                html.append(u'<details><summary>{} — x{}</summary>'.format(_h(tname), len(items)))
                rows=[]
                for it in items:
                    link = out.linkify(DB.ElementId(it["id"]), u"{}".format(it["id"])) if it.get("id") else u""
                    rows.append([link, _h(it.get("cat") or u""), _h(it.get("reason") or u"")])
                html.append(_table([u"ID", u"Категория", u"Причина"], rows,
                                   align=["right","left","left"], safe=[0]))
                html.append(u'</details>')
            html.append(u'</details>')

    html.append(u"</div></div>")
    out.print_html(u"".join(html))

# ---- XLSX (минимальный OpenXML, на случай проверки) ----
def _xlsx_cell(v, is_text=False):
    if v is None or v == "": return u'<c/>'
    if is_text:
        return u'<c t="inlineStr"><is><t>{}</t></is></c>'.format(_h(v))
    try:
        return u'<c><v>{:.6f}</v></c>'.format(float(v))
    except:
        return u'<c t="inlineStr"><is><t>{}</t></is></c>'.format(_h(v))

def _xlsx_sheet_xml(headers, rows, text_cols=None):
    text_cols = set(text_cols or [])
    lines = [u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    lines.append(u'<row r="1">')
    for h in headers: lines.append(_xlsx_cell(h, True))
    lines.append(u'</row>')
    r=2
    for row in rows:
        lines.append(u'<row r="{}">'.format(r))
        for c,val in enumerate(row):
            lines.append(_xlsx_cell(val, is_text=(c in text_cols)))
        lines.append(u'</row>'); r+=1
    lines.append(u'</sheetData></worksheet>')
    return u"".join(lines)

def _xlsx_build(filepath, calc_map, skip_map, totals):
    # лист 1 — Summary totals
    sum_headers = [u"Метрика", u"Значение"]
    sum_rows = [
        [u"Нормативная стоимость (ГЭСН)", _fmt_money(totals["N"])],
        [u"Опытная стоимость",            _fmt_money(totals["F"])],
        [u"Нормативные трудозатраты",     _fmt_num(totals["LN"])],
        [u"Опытные трудозатраты",         _fmt_num(totals["LF"])],
    ]
    # лист 2 — Details (по элементам рассчитанным)
    det_headers = [u"ID", u"Стадия", u"Тип", u"Категория", u"ЕИ", u"Кол-во", u"Н цена/ед", u"Ф цена/ед",
                   u"Н стоимость", u"Ф стоимость", u"Н труд.", u"Ф труд."]
    det_rows=[]
    for stage in calc_map:
        for tname, data in calc_map[stage].items():
            for it in data["items"]:
                det_rows.append([it["id"], stage, tname, it["cat"], it["unit"], it["qty"],
                                 it["rcn"], it["rcf"], it["cn"], it["cf"], it["ln"], it["lf"]])
    summary_xml = _xlsx_sheet_xml(sum_headers, sum_rows, text_cols=set([0]))
    details_xml = _xlsx_sheet_xml(det_headers, det_rows, text_cols=set([1,2,3,4]))

    # упаковка
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    content_types = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"  ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels_root = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    wb_rels = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>'''
    workbook = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Details" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>'''
    core = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ACBD Calculation Export</dc:title>
  <dc:creator>pyRevit</dc:creator>
  <cp:lastModifiedBy>pyRevit</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">%(now)s</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">%(now)s</dcterms:modified>
</cp:coreProperties>''' % {"now": now}
    app = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/docPropsVTypes">
  <Application>pyRevit</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>'''

    z = zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED)
    try:
        z.writestr("[Content_Types].xml",        content_types.encode("utf-8"))
        z.writestr("_rels/.rels",                 rels_root.encode("utf-8"))
        z.writestr("docProps/core.xml",          core.encode("utf-8"))
        z.writestr("docProps/app.xml",           app.encode("utf-8"))
        z.writestr("xl/workbook.xml",            workbook.encode("utf-8"))
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels.encode("utf-8"))
        z.writestr("xl/worksheets/sheet1.xml",   summary_xml.encode("utf-8"))
        z.writestr("xl/worksheets/sheet2.xml",   details_xml.encode("utf-8"))
    finally:
        z.close()

# ---- Выбор области и режима ----
class _ScopeDialog(object):
    def __init__(self, default_visible=False):
        self._result = None

        wnd = Window()
        wnd.Title = u"ACBD"
        wnd.Width = 260
        wnd.Height = 180
        wnd.ResizeMode = ResizeMode.NoResize
        wnd.WindowStyle = WindowStyle.ToolWindow
        try:
            from System.Windows import WindowStartupLocation  # noqa: WPS433
            wnd.WindowStartupLocation = WindowStartupLocation.CenterOwner
        except Exception:
            pass

        stack = StackPanel()
        stack.Margin = Thickness(12)

        label = TextBlock()
        label.Text = u"Что пересчитывать?"
        label.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(label)

        self._scope_all = RadioButton()
        self._scope_all.Content = u"Вся модель"
        self._scope_all.Margin = Thickness(0, 0, 0, 4)
        self._scope_all.IsChecked = not default_visible
        stack.Children.Add(self._scope_all)

        self._scope_visible = RadioButton()
        self._scope_visible.Content = u"Видимые элементы"
        self._scope_visible.IsChecked = default_visible
        stack.Children.Add(self._scope_visible)

        self._recon = CheckBox()
        self._recon.Content = u"Реконструкция"
        self._recon.Margin = Thickness(0, 12, 0, 0)
        stack.Children.Add(self._recon)

        buttons = StackPanel()
        buttons.Orientation = Orientation.Horizontal
        buttons.HorizontalAlignment = HorizontalAlignment.Right
        buttons.Margin = Thickness(0, 16, 0, 0)

        ok_btn = Button()
        ok_btn.Content = u"OK"
        ok_btn.Width = 80
        ok_btn.Margin = Thickness(0, 0, 6, 0)
        ok_btn.IsDefault = True
        ok_btn.Click += self._on_ok
        buttons.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = u"Отмена"
        cancel_btn.Width = 80
        cancel_btn.IsCancel = True
        cancel_btn.Click += self._on_cancel
        buttons.Children.Add(cancel_btn)

        stack.Children.Add(buttons)

        wnd.Content = stack
        self._window = wnd

    def _on_ok(self, sender, args):
        scope = u"Видимые элементы" if bool(self._scope_visible.IsChecked) else u"Вся модель"
        recon = bool(self._recon.IsChecked)
        self._result = (scope, recon)
        try:
            self._window.DialogResult = True
        except Exception:
            pass
        self._window.Close()

    def _on_cancel(self, sender, args):
        self._result = None
        try:
            self._window.DialogResult = False
        except Exception:
            pass
        self._window.Close()

    def show_dialog(self):
        try:
            self._window.ShowDialog()
        except Exception:
            self._window.Show()
        return self._result


def _select_scope(default_visible=False):
    dlg = _ScopeDialog(default_visible=default_visible)
    result = dlg.show_dialog()
    if not result:
        script.exit()
    return result


# ---- Запуск ----
choice, reconstruction_mode = _select_scope(default_visible=False)
elements = _collect_visible(revit.active_view) if choice == u"Видимые элементы" else _collect_all()
if reconstruction_mode:
    allowed_stages = {ST_DEMOL, ST_NEW}
    elements = [el for el in elements if _stage_bucket(el) in allowed_stages]

totals = dict(N=0.0, F=0.0, LN=0.0, LF=0.0)
calc_map = {}   # stage -> type -> {sumN,sumF,sumLN,sumLF,count,items[]}
skip_map = {}   # stage -> type -> [ {id,cat,tname,reason} ]

okcnt = 0
with revit.Transaction(u"ACBD: пересчёт стоимости и трудозатрат"):
    for el in elements:
        if _calc_element(el, calc_map, skip_map, totals):
            okcnt += 1

_render_report(calc_map, skip_map, totals, len(elements), okcnt)

# Предлагаем сохранить XLSX (опционально)
fname = u"ACBD_Calc_{:%Y%m%d_%H%M}.xlsx".format(datetime.datetime.now())
save = forms.save_file(file_ext="xlsx", default_name=fname, title=u"Сохранить отчёт XLSX (по рассчитанным)")
if save:
    try:
        _xlsx_build(save, calc_map, skip_map, totals)
        out.print_html(u'<p><b>XLSX сохранён:</b> <span class="mono">{}</span></p>'.format(_h(save)))
    except Exception as e:
        out.print_html(u'<p><b>Ошибка записи XLSX:</b> {}</p>'.format(_h(e)))
