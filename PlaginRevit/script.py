# -*- coding: utf-8 -*-
import re, os, ctypes
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan

doc = revit.doc
out = script.get_output()

# -------------------- параметры --------------------
P_COST_N_RATE = u"ACBD_Н_ЦенаЗаЕдИзм"
P_COST_F_RATE = u"ACBD_Ф_ЦенаЗаЕдИзм"
P_LAB_N_RATE  = u"ACBD_Н_ТрудозатратыНаЕдИзм"
P_LAB_F_RATE  = u"ACBD_Ф_ТрудозатратыНаЕдИзм"

P_COST_N = u"ACBD_Н_СтоимостьЭлемента"
P_COST_F = u"ACBD_Ф_СтоимостьЭлемента"
P_LAB_N  = u"ACBD_Н_ТрудозатратыЭлемента"
P_LAB_F  = u"ACBD_Ф_ТрудозатратыЭлемента"

UNIT_PARAM_NAMES = (
    u"ACBD_Н_ЕдиницаИзмерения",
    u"ACBD_ЕдиницаИзмерения",
    u"ADSK_Единица измерения",
    u"Единица измерения",
    u"Ед. изм.", u"Ед. изм",
    u"Unit",
)

# -------------------- утилиты --------------------
try:
    _text_type = unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover - python 3
    _text_type = str


def _to_text(v):
    if v is None:
        return None
    try:
        if isinstance(v, _text_type):
            return _text_type(v)
    except Exception:
        pass
    to_string = getattr(v, "ToString", None)
    if callable(to_string):
        try:
            return _text_type(to_string())
        except Exception:
            pass
    try:
        return _text_type(v)
    except Exception:
        pass
    try:
        return _text_type(str(v))
    except Exception:
        pass
    return None

def _num(v):
    if v is None: return None
    s = _to_text(v)
    if not s: return None
    s = re.sub(u"[^0-9,.-]", u"", s.strip()).replace(u",", u".")
    try: return float(s)
    except: return None

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

# -------------------- отбор строительных элементов --------------------
ALLOWED = List[DB.BuiltInCategory]([
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

def _has_any_acbd(el):
    names = (P_COST_N_RATE,P_COST_F_RATE,P_LAB_N_RATE,P_LAB_F_RATE,P_COST_N,P_COST_F,P_LAB_N,P_LAB_F)
    for n in names:
        if _get_param(el, n): return True
    t = _get_type(el)
    if t:
        for n in names:
            if _get_param(t, n): return True
    return False

def _elements_all():
    f_cats = DB.ElementMulticategoryFilter(ALLOWED)
    col = (DB.FilteredElementCollector(doc)
             .WhereElementIsNotElementType()
             .WherePasses(f_cats))
    res = []
    for el in col:
        if getattr(el, "ViewSpecific", False):  # детализ. элементы и т.п.
            continue
        cat = getattr(el, "Category", None)
        if not cat or cat.CategoryType != DB.CategoryType.Model:
            continue
        if not _has_any_acbd(el):
            continue
        res.append(el)
    return res

def _elements_visible(view):
    """Элементы, видимые на активном виде."""
    f_cats = DB.ElementMulticategoryFilter(ALLOWED)
    col = (DB.FilteredElementCollector(doc, view.Id)
             .WhereElementIsNotElementType()
             .WherePasses(f_cats))
    res = []
    for el in col:
        cat = getattr(el, "Category", None)
        if not cat or cat.CategoryType != DB.CategoryType.Model:
            continue
        if not _has_any_acbd(el):
            continue
        res.append(el)
    return res

# -------------------- количества --------------------
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

# -------------------- запись чисел --------------------
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
    p = _get_param(el, target_name)
    if p and not getattr(p, "IsReadOnly", False):
        if _try_set_with_formats(p, value, _is_currency_param(p)):
            return True, "instance"
    t = _get_type(el)
    if t:
        pt = _get_param(t, target_name)
        if pt and not getattr(pt, "IsReadOnly", False):
            if _try_set_with_formats(pt, value, _is_currency_param(pt)):
                return True, "type"
    return False, "missing"

# -------------------- счётчики/итоги --------------------
w_inst = w_type = 0
auto_units = 0
cost_written = labor_written = 0
total_cost_n = 0.0
total_cost_f = 0.0

def _apply_value(el, target_name, value, is_cost):
    global w_inst, w_type, cost_written, labor_written
    ok, where = _set_number_any(el, target_name, value)
    if ok:
        if where == "instance": w_inst += 1
        else: w_type += 1
        if is_cost: cost_written += 1
        else: labor_written += 1
    return ok

def _calc_and_set(el):
    global auto_units, total_cost_n, total_cost_f
    unit_txt, qty, src = _get_unit_and_qty(el)
    if src != u"param": auto_units += 1

    r_cn = _get_rate(el, P_COST_N_RATE)
    r_cf = _get_rate(el, P_COST_F_RATE)
    r_ln = _get_rate(el, P_LAB_N_RATE)
    r_lf = _get_rate(el, P_LAB_F_RATE)

    if r_cn is not None:
        val = (r_cn or 0.0) * (qty or 0.0)
        total_cost_n += val
        _apply_value(el, P_COST_N, val, True)
    if r_cf is not None:
        val = (r_cf or 0.0) * (qty or 0.0)
        total_cost_f += val
        _apply_value(el, P_COST_F, val, True)
    if r_ln is not None:
        _apply_value(el, P_LAB_N, (r_ln or 0.0) * (qty or 0.0), False)
    if r_lf is not None:
        _apply_value(el, P_LAB_F, (r_lf or 0.0) * (qty or 0.0), False)

# -------------------- оверлей-окно (безопасное) --------------------
class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

def _get_revit_rect():
    try:
        hwnd = revit.uidoc.Application.MainWindowHandle
        r = RECT()
        ctypes.windll.user32.GetWindowRect(ctypes.c_void_p(int(hwnd.ToInt64())), ctypes.byref(r))
        return r.left, r.top, r.right, r.bottom
    except:
        return None

def _ensure_overlay_xaml(path):
    if os.path.exists(path): return
    xaml = u'''<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Стоимость проекта" Width="360" Height="140"
    WindowStartupLocation="Manual" Left="10" Top="80"
    WindowStyle="None" AllowsTransparency="True" Background="Transparent"
    ResizeMode="NoResize" Topmost="True" ShowInTaskbar="False">
      <Border CornerRadius="10" Background="#EE202020" Padding="10" x:Name="DragArea">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <DockPanel>
            <TextBlock Text="Стоимость проектируемого объекта"
                       Foreground="White" FontWeight="Bold" FontSize="14" DockPanel.Dock="Left"/>
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
              <ToggleButton x:Name="PinBtn" Content="📌" Width="26" Height="24"
                            Margin="0,0,6,0" Background="#33FFFFFF" Foreground="White" BorderThickness="0"
                            ToolTip="Фиксировать к углу/Свободно" IsChecked="True"/>
              <Button x:Name="CloseBtn" Content="×" Width="26" Height="24"
                      Background="#33FFFFFF" Foreground="White" BorderThickness="0"/>
            </StackPanel>
          </DockPanel>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,8,0,0">
            <TextBlock Text="Н: " Foreground="White" FontSize="16"/>
            <TextBlock x:Name="CostN" Foreground="White" FontSize="16" FontWeight="Bold"/>
          </StackPanel>
          <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,4,0,0">
            <TextBlock Text="Ф: " Foreground="White" FontSize="16"/>
            <TextBlock x:Name="CostF" Foreground="White" FontSize="16" FontWeight="Bold"/>
          </StackPanel>
        </Grid>
      </Border>
    </Window>'''
    f = open(path, "wb"); f.write(xaml.encode("utf-8")); f.close()

def _fmt_rub(val):
    try: return (u"{:,.2f} ₽".format(float(val))).replace(u",", u" ").replace(u".", u",")
    except: return _text_type(val)

def _show_overlay(total_n, total_f):
    sticky = script.get_sticky(); KEY = "ACBD_COST_OVERLAY"

    old = sticky.get(KEY)
    if old:
        try: old["timer"].Stop()
        except: pass
        try: old["wnd"].Close()
        except: pass
        sticky.pop(KEY, None)

    xaml_path = os.path.join(os.path.dirname(__file__), "ProjectCostOverlay.xaml")
    _ensure_overlay_xaml(xaml_path)
    wnd = forms.WPFWindow(xaml_path)
    try:
        hwnd = revit.uidoc.Application.MainWindowHandle
        WindowInteropHelper(wnd).Owner = hwnd
    except: pass
    try:
        wnd.CostN.Text = _fmt_rub(total_n)
        wnd.CostF.Text = _fmt_rub(total_f)
    except: pass

    state = {"corner":"TL", "dx":10, "dy":80}

    def apply_anchor():
        rc = _get_revit_rect()
        if not rc: return
        L,T,R,B = rc
        if state["corner"] == "TL":
            wnd.Left = L + state["dx"]; wnd.Top = T + state["dy"]
        elif state["corner"] == "TR":
            wnd.Left = R - wnd.Width - state["dx"]; wnd.Top = T + state["dy"]
        elif state["corner"] == "BL":
            wnd.Left = L + state["dx"]; wnd.Top = B - wnd.Height - state["dy"]
        else:
            wnd.Left = R - wnd.Width - state["dx"]; wnd.Top = B - wnd.Height - state["dy"]

    def pick_nearest_corner():
        rc = _get_revit_rect()
        if not rc: return
        L,T,R,B = rc
        cx = wnd.Left + wnd.Width/2.0; cy = wnd.Top + wnd.Height/2.0
        corners = {"TL":(L,T), "TR":(R,T), "BL":(L,B), "BR":(R,B)}
        best, dmin = "TL", 10**9
        for k,(x,y) in corners.items():
            d = (cx-x)*(cx-x)+(cy-y)*(cy-y)
            if d < dmin: best, dmin = k, d
        state["corner"] = best
        if best == "TL":
            state["dx"] = max(8, int(wnd.Left - L));             state["dy"] = max(8, int(wnd.Top - T))
        elif best == "TR":
            state["dx"] = max(8, int(R - (wnd.Left + wnd.Width))); state["dy"] = max(8, int(wnd.Top - T))
        elif best == "BL":
            state["dx"] = max(8, int(wnd.Left - L));             state["dy"] = max(8, int(B - (wnd.Top + wnd.Height)))
        else:
            state["dx"] = max(8, int(R - (wnd.Left + wnd.Width))); state["dy"] = max(8, int(B - (wnd.Top + wnd.Height)))
        apply_anchor()

    def on_drag(sender, args):
        try:
            wnd.DragMove()
            if wnd.PinBtn.IsChecked:
                pick_nearest_corner()
        except: pass
    try: wnd.DragArea.MouseLeftButtonDown += on_drag
    except: pass

    timer = DispatcherTimer()
    timer.Interval = TimeSpan.FromMilliseconds(400)
    def on_tick(sender, args):
        try:
            if wnd.PinBtn.IsChecked:
                apply_anchor()
        except: pass
    timer.Tick += on_tick

    def on_close(sender, args):
        try: timer.Stop()
        except: pass
        sticky.pop(KEY, None)
    try: wnd.CloseBtn.Click += on_close
    except: pass
    try: wnd.Closed += on_close
    except: pass

    apply_anchor()
    timer.Start()
    sticky[KEY] = {"wnd": wnd, "timer": timer}
    try: wnd.show()
    except: wnd.show_dialog()

# -------------------- запуск --------------------

# диалог выбора области
choice = forms.CommandSwitchWindow.show(
    [u"Вся модель", u"Видимые элементы"],
    message=u"Что пересчитывать?"
) or u"Вся модель"

if choice == u"Видимые элементы":
    elements = _elements_visible(revit.active_view)
else:
    elements = _elements_all()

with revit.Transaction(u"ACBD: расчёт стоимости/трудозатрат"):
    for el in elements:
        _calc_and_set(el)

_show_overlay(total_cost_n, total_cost_f)

out.print_md(u"### Готово")
out.print_md(u"Элементов: **{0}** | В экземпляры: **{1}** | В типы: **{2}**"
             .format(len(elements), w_inst, w_type))
out.print_md(u"Стоимость записана: **{0}**, Трудозатраты записаны: **{1}**"
             .format(cost_written, labor_written))
out.print_md(u"ИТОГО Н: **{0}**, ИТОГО Ф: **{1}**"
             .format(_fmt_currency(total_cost_n), _fmt_currency(total_cost_f)))
