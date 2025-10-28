# -*- coding: utf-8 -*-

import re
from pyrevit import revit, DB, script
from System.Collections.Generic import List as CsList
from System.Windows import (Application, Window, WindowStyle, ResizeMode, Thickness, FontWeights,
                            HorizontalAlignment)
from System.Windows.Controls import (Border, StackPanel, TextBlock, Orientation, Separator,
                                     RadioButton, CheckBox, Button)
from System.Windows.Media import SolidColorBrush, Color, Brushes

doc = revit.doc

# --------- параметры ACBD ---------
P_UNIT_T    = u"ACBD_ЕдиницаИзмерения"
P_RATE_CN_T = u"ACBD_Н_ЦенаЗаЕдИзм"
P_RATE_CF_T = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_RATE_LN_T = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_RATE_LF_T = u"ACBD_Ф_ТрудозатратыНаЕдИзм"

P_COST_N_I  = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F_I  = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N_I   = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F_I   = u"ACBD_Ф_ТрудозатратыЭлемента"

# --------- helpers ---------
try:
    text_type = unicode
except NameError:
    text_type = str

_C2L = {u"А":u"A",u"а":u"a",u"В":u"B",u"в":u"b",u"С":u"C",u"с":u"c",u"Е":u"E",u"е":u"e",
        u"Н":u"H",u"н":u"h",u"К":u"K",u"к":u"k",u"М":u"M",u"м":u"m",u"О":u"O",u"о":u"o",
        u"Р":u"P",u"р":u"p",u"Т":u"T",u"т":u"t",u"Х":u"X",u"х":u"x",u"У":u"Y",u"у":u"y"}

def _t(x):
    if x is None: return None
    try:
        if isinstance(x, text_type): return x
    except: pass
    try: return text_type(x)
    except:
        try: return text_type(x.ToString())
        except: return None

def _fold_name(s):
    s = (_t(s) or u"").replace(u"\u00A0", u" ").strip()
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

def _fmt_labor(v):
    try:
        return (u"{:,.2f}".format(float(v))).replace(u",", u" ").replace(u".", u",")
    except:
        return _t(v) or u""

def _lp(holder, name):
    if not holder: return None
    try:
        p = holder.LookupParameter(name)
        if p: return p
    except: pass
    want = _fold_name(name)
    try:
        for p in holder.Parameters:
            dn = _fold_name(getattr(p.Definition, "Name", u""))
            if dn == want: return p
    except: pass
    return None

def _eltype(el):
    try: return doc.GetElement(el.GetTypeId())
    except: return None

def _get_str_from(holder, name):
    p = _lp(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.String:
            return _t(p.AsString())
        return _t(p.AsValueString())
    except: return None

def _get_num_from(holder, name):
    p = _lp(holder, name)
    if not p: return None
    try:
        if p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
        if p.StorageType == DB.StorageType.String:
            return _num(p.AsString())
        return _num(p.AsValueString())
    except: return None

def _inst_param(el, name):
    return _lp(el, name)

def _is_currency(p):
    try:
        dt = p.Definition.GetDataType()
        if dt and dt.Equals(DB.SpecTypeId.Currency): return True
    except: pass
    try:
        return getattr(p.Definition, "ParameterType", None) == DB.ParameterType.Currency
    except: return False

def _try_set_number(p, value):
    try:
        if p.Set(float(value)): return True
    except: pass
    variants = (
        _fmt_money(value),
        u"{:.2f}".format(float(value)).replace(u".", u","),
        u"{:,.2f}".format(float(value)).replace(u",", u" ").replace(u".", u","),
        u"{:.2f}".format(float(value)),
    ) if _is_currency(p) else (
        u"{:,.2f}".format(float(value)).replace(u",", u" ").replace(u".", u","),
        u"{:.2f}".format(float(value)),
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

# --------- отбор строительных элементов ---------
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

ST_EXIST = u"Существующие"
ST_DEMOL = u"Демонтаж"
ST_NEW   = u"Новые конструкции"
ST_OTHER = u"Прочее"


def _phase_names(el):
    cr = None
    dm = None
    try:
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if p:
            pid = p.AsElementId()
            if pid and pid.IntegerValue > 0:
                ph = doc.GetElement(pid)
                cr = _t(getattr(ph, "Name", None))
    except:  # noqa: E722 - API access safety
        pass
    try:
        p = el.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
        if p:
            pid = p.AsElementId()
            if pid and pid.IntegerValue > 0:
                ph = doc.GetElement(pid)
                dm = _t(getattr(ph, "Name", None))
    except:  # noqa: E722 - API access safety
        pass
    return cr, dm


def _nru(s):
    return (_t(s) or u"").strip().lower().replace(u"\u00a0", u" ")


def _stage_bucket(el):
    cr, dm = _phase_names(el)
    crl, dml = _nru(cr), _nru(dm)
    if (dml == u"демонтаж") and (crl == u"существующие"):
        return ST_DEMOL
    if (not dml) and (crl == u"новая конструкция"):
        return ST_NEW
    if crl == u"существующие":
        return ST_EXIST
    if u"демонтаж" in dml and u"существ" in crl:
        return ST_DEMOL
    if (not dml) and (u"нов" in crl):
        return ST_NEW
    if u"существ" in crl:
        return ST_EXIST
    return ST_OTHER

def _collect_all():
    f = DB.ElementMulticategoryFilter(ALLOWED)
    col = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(f)
    return [el for el in col
            if getattr(el, "ViewSpecific", False) is False
            and getattr(getattr(el, "Category", None), "CategoryType", None) == DB.CategoryType.Model]

def _collect_visible(view):
    f = DB.ElementMulticategoryFilter(ALLOWED)
    vis = DB.VisibleInViewFilter(doc, view.Id)
    col = (DB.FilteredElementCollector(doc, view.Id)
           .WhereElementIsNotElementType()
           .WherePasses(f)
           .WherePasses(vis))
    return [el for el in col
            if getattr(getattr(el, "Category", None), "CategoryType", None) == DB.CategoryType.Model]

# --------- количества из ИНСТАНС-параметров ---------
def _get_double_si(el, names, unit_tid):
    if not isinstance(names, (list, tuple)): names = (names,)
    for nm in names:
        p = _inst_param(el, nm)
        if not p: continue
        try:
            if p.StorageType == DB.StorageType.Double:
                return DB.UnitUtils.ConvertFromInternalUnits(p.AsDouble(), unit_tid)
            if p.StorageType == DB.StorageType.String:
                v = _num(p.AsString())
                if v is not None: return v
            vs = _t(p.AsValueString()); v = _num(vs)
            if v is not None: return v
        except: pass
    return None

def _qty(el, unit_text):
    if not unit_text: return None
    key = (_t(unit_text) or u"").strip().lower()
    key = key.replace(u"\u00a0", u" ").replace(u" ", u"")
    if key in (u"квм", u"кв.м", u"м2", u"м²", u"m2", u"sqm"): key = u"м2"
    if key in (u"кубм", u"куб.м", u"м3", u"м³", u"m3", u"cbm"): key = u"м3"
    if key in (u"м", u"мп", u"м.п", u"м.п.", u"п.м", u"pm", u"rm"): key = u"м"
    if key in (u"шт", u"шт.", u"штука", u"pcs", u"pc"): key = u"шт"

    if key == u"м2":
        v = _get_double_si(el, (u"Area", u"Площадь"), DB.UnitTypeId.SquareMeters)
        return 0.0 if v is None else v
    if key == u"м3":
        v = _get_double_si(el, (u"Volume", u"Объем", u"Объём"), DB.UnitTypeId.CubicMeters)
        return 0.0 if v is None else v
    if key == u"м":
        v = _get_double_si(el, (u"Length", u"Длина", u"Perimeter"), DB.UnitTypeId.Meters)
        return 0.0 if v is None else v
    if key == u"шт":
        return 1.0
    return None

# --------- расчёт одного элемента ---------
def _calc_element(el):
    et = _eltype(el)
    unit_text = _get_str_from(et, P_UNIT_T)
    if not unit_text or not _t(unit_text).strip():
        return False, None, None, None, None

    q = _qty(el, unit_text)
    if q is None:
        return False, None, None, None, None

    r_cn = _get_num_from(et, P_RATE_CN_T)
    r_cf = _get_num_from(et, P_RATE_CF_T)
    r_ln = _get_num_from(et, P_RATE_LN_T)
    r_lf = _get_num_from(et, P_RATE_LF_T)

    ok_any = False
    cost_n = cost_f = lab_n = lab_f = None

    if r_cn is not None:
        cost_n = (r_cn or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_COST_N_I, cost_n)
    if r_cf is not None:
        cost_f = (r_cf or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_COST_F_I, cost_f)
    if r_ln is not None:
        lab_n = (r_ln or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_LAB_N_I, lab_n)
    if r_lf is not None:
        lab_f = (r_lf or 0.0) * (q or 0.0)
        ok_any |= _set_inst_number(el, P_LAB_F_I, lab_f)

    return ok_any, cost_n, cost_f, lab_n, lab_f

# --------- одно окно "Стоимость объекта" с сноской ---------
COST_TAG   = "ACBD_COST_WINDOW"
VALN_TAG   = "ACBD_VALN"    # Стоимость Н
VALF_TAG   = "ACBD_VALF"    # Стоимость Ф
VALLN_TAG  = "ACBD_VALLN"   # Трудозатраты Н
VALLF_TAG  = "ACBD_VALLF"   # Трудозатраты Ф
FOOTER_TAG = "ACBD_FOOTER"  # Сноска

def _walk(o):
    if o is None: return
    yield o
    for attr in ('Child', 'Content'):
        try:
            ch = getattr(o, attr, None)
            if ch and ch is not o:
                for x in _walk(ch): yield x
        except: pass
    try:
        chs = getattr(o, 'Children', None)
        if chs:
            for c in chs:
                for x in _walk(c): yield x
    except: pass

def _find_child_by_tag(root, tag):
    for n in _walk(root):
        try:
            if getattr(n, 'Tag', None) == tag:
                return n
        except: pass
    return None

def _find_window(tag):
    try:
        app = Application.Current
        if not app: return None
        found = []
        for w in app.Windows:
            try:
                if getattr(w, "Tag", None) == tag and getattr(w, "IsVisible", False):
                    found.append(w)
            except: pass
        if len(found) > 1:
            for w in found[1:]:
                try: w.Close()
                except: pass
        return found[0] if found else None
    except:
        return None

def _build_cost_content(wnd):
    border = Border(); border.Padding = Thickness(10)
    try: border.Background = SolidColorBrush(Color.FromRgb(32,32,32))
    except: pass
    wnd.Content = border

    stack = StackPanel(); border.Child = stack

    title = TextBlock()
    title.Text = u"Стоимость проектируемого объекта"
    title.FontWeight = FontWeights.Bold; title.FontSize = 14
    try: title.Foreground = Brushes.White
    except: pass
    stack.Children.Add(title)

    capCost = TextBlock(); capCost.Text = u"СТОИМОСТЬ СТРОИТЕЛЬСТВА"; capCost.Margin = Thickness(0,8,0,2)
    capCost.FontWeight = FontWeights.Bold
    try: capCost.Foreground = Brushes.White
    except: pass
    stack.Children.Add(capCost)

    row1 = StackPanel(); row1.Orientation = Orientation.Horizontal
    t1 = TextBlock(); t1.Text = u"Нормативная: "; t1.FontSize = 16
    v1 = TextBlock(); v1.FontSize = 16; v1.FontWeight = FontWeights.Bold; v1.Tag = VALN_TAG
    try: t1.Foreground = Brushes.White; v1.Foreground = Brushes.White
    except: pass
    row1.Children.Add(t1); row1.Children.Add(v1); stack.Children.Add(row1)

    row2 = StackPanel(); row2.Orientation = Orientation.Horizontal; row2.Margin = Thickness(0,2,0,0)
    t2 = TextBlock(); t2.Text = u"Опытная: "; t2.FontSize = 16
    v2 = TextBlock(); v2.FontSize = 16; v2.FontWeight = FontWeights.Bold; v2.Tag = VALF_TAG
    try: t2.Foreground = Brushes.White; v2.Foreground = Brushes.White
    except: pass
    row2.Children.Add(t2); row2.Children.Add(v2); stack.Children.Add(row2)

    capLab = TextBlock(); capLab.Text = u"ТРУДОЗАТРАТЫ"; capLab.Margin = Thickness(0,10,0,2)
    capLab.FontWeight = FontWeights.Bold
    try: capLab.Foreground = Brushes.White
    except: pass
    stack.Children.Add(capLab)

    row3 = StackPanel(); row3.Orientation = Orientation.Horizontal
    t3 = TextBlock(); t3.Text = u"Нормативная: "; t3.FontSize = 16
    v3 = TextBlock(); v3.FontSize = 16; v3.FontWeight = FontWeights.Bold; v3.Tag = VALLN_TAG
    try: t3.Foreground = Brushes.White; v3.Foreground = Brushes.White
    except: pass
    row3.Children.Add(t3); row3.Children.Add(v3); stack.Children.Add(row3)

    row4 = StackPanel(); row4.Orientation = Orientation.Horizontal; row4.Margin = Thickness(0,2,0,0)
    t4 = TextBlock(); t4.Text = u"Опытная: "; t4.FontSize = 16
    v4 = TextBlock(); v4.FontSize = 16; v4.FontWeight = FontWeights.Bold; v4.Tag = VALLF_TAG
    try: t4.Foreground = Brushes.White; v4.Foreground = Brushes.White
    except: pass
    row4.Children.Add(t4); row4.Children.Add(v4); stack.Children.Add(row4)

    sep = Separator(); sep.Margin = Thickness(0,10,0,6)
    stack.Children.Add(sep)

    foot = TextBlock()
    foot.Tag = FOOTER_TAG
    foot.FontSize = 11
    try: foot.Foreground = Brushes.Gainsboro
    except: pass
    stack.Children.Add(foot)

def _ensure_cost_window():
    wnd = _find_window(COST_TAG)
    if not wnd:
        wnd = Window()
        wnd.Title = u"Стоимость объекта"
        wnd.Width = 440; wnd.Height = 260
        wnd.WindowStartupLocation = 0; wnd.Left = 20; wnd.Top = 80
        wnd.WindowStyle = WindowStyle.ToolWindow
        wnd.Topmost = True; wnd.ResizeMode = ResizeMode.NoResize
        wnd.ShowInTaskbar = False; wnd.Tag = COST_TAG
        _build_cost_content(wnd)
        try: wnd.Show()
        except: wnd.ShowDialog()
    else:
        need = [VALN_TAG, VALF_TAG, VALLN_TAG, VALLF_TAG, FOOTER_TAG]
        if any(_find_child_by_tag(wnd, tag) is None for tag in need):
            _build_cost_content(wnd)
    return wnd

def _update_cost_window(total_n, total_f, total_ln, total_lf, processed, okcnt, skipped, scope_text):
    def fmt_money(v):
        try:  return (u"{:,.2f} ₽".format(float(v))).replace(u",", u" ").replace(u".", u",")
        except: return u"{}".format(v)
    def fmt_lab(v):
        return _fmt_labor(v) + u" ч*ч" if v is not None else u"—"
    wnd = _ensure_cost_window()
    v1 = _find_child_by_tag(wnd, VALN_TAG)
    v2 = _find_child_by_tag(wnd, VALF_TAG)
    v3 = _find_child_by_tag(wnd, VALLN_TAG)
    v4 = _find_child_by_tag(wnd, VALLF_TAG)
    if v1: v1.Text = fmt_money(total_n)
    if v2: v2.Text = fmt_money(total_f)
    if v3: v3.Text = fmt_lab(total_ln)
    if v4: v4.Text = fmt_lab(total_lf)
    foot = _find_child_by_tag(wnd, FOOTER_TAG)
    if foot:
        foot.Text = (u"Обработано: {0} | С расчётом: {1} | Пропущено: {2} | Область: {3}"
                     .format(processed, okcnt, skipped, _t(scope_text) or u"—"))
    try:
        if not wnd.IsVisible: wnd.Show()
    except: pass

# --------- выбор области и режима ---------


class _ScopeDialog(object):
    def __init__(self, default_visible=True):
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
        self._scope_all.IsChecked = bool(not default_visible)
        stack.Children.Add(self._scope_all)

        self._scope_visible = RadioButton()
        self._scope_visible.Content = u"Видимые элементы"
        self._scope_visible.IsChecked = bool(default_visible)
        stack.Children.Add(self._scope_visible)

        self._recon = CheckBox()
        self._recon.Content = u"Реконструкция"
        self._recon.Margin = Thickness(0, 12, 0, 0)
        self._recon.IsChecked = False
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
        scope = u"Видимые элементы" if self._scope_visible.IsChecked is True else u"Вся модель"
        recon = (self._recon.IsChecked is True)
        self._result = (scope, recon)
        try:
            self._window.DialogResult = True
        except Exception:  # noqa: WPS466
            pass
        self._window.Close()

    def _on_cancel(self, sender, args):
        self._result = None
        try:
            self._window.DialogResult = False
        except Exception:  # noqa: WPS466
            pass
        self._window.Close()

    def show_dialog(self):
        try:
            self._window.ShowDialog()
        except Exception:  # noqa: WPS466 - fallback for modeless environments
            self._window.Show()
        return self._result


def _select_scope(default_visible=True):
    dlg = _ScopeDialog(default_visible=default_visible)
    result = dlg.show_dialog()
    if not result:
        script.exit()
    return result


# --------- запуск: выбор области + расчёт ---------
choice, reconstruction_mode = _select_scope(default_visible=True)
scope_text = choice + (u"; реконструкция" if reconstruction_mode else u"")

elements = _collect_visible(revit.active_view) if choice == u"Видимые элементы" else _collect_all()
if reconstruction_mode:
    allowed_stages = {ST_DEMOL, ST_NEW}
    elements = [el for el in elements if _stage_bucket(el) in allowed_stages]

total_n = 0.0
total_f = 0.0
total_ln = 0.0
total_lf = 0.0
ok_count = 0
skipped_count = 0

with revit.Transaction(u"ACBD: расчёт стоимости и трудозатрат"):
    for el in elements:
        ok, cN, cF, lN, lF = _calc_element(el)
        if ok:
            ok_count += 1
            if cN is not None: total_n += float(cN)
            if cF is not None: total_f += float(cF)
            if lN is not None: total_ln += float(lN)
            if lF is not None: total_lf += float(lF)
        else:
            skipped_count += 1

_update_cost_window(total_n, total_f, total_ln, total_lf,
                    len(elements), ok_count, skipped_count, scope_text)
