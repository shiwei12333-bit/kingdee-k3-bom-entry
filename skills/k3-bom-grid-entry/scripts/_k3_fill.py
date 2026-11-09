# -*- coding: utf-8 -*-
"""
金蝶 K3「BOM单 - 新增」明细表录入器（带自动锚定，防自动滚动错行）

核心思想（v2 —— 经循环自检修正）：
  1) 不用固定 row_y(n)，每次先用 click 命中「商品代码」列拿到 focused().rect.top 做行锚；
  2) ⚠️ automation_id 200/201 **不是列唯一**：多个文本列都是 200、多个数值列（用量/损耗率）
     都是 201。判定命中哪一行哪一列必须同时看 (aid, rect.left)，光看 aid 会静默写错列；
  3) 只填 商品代码 + 用量 两列，其余列 K3 自动带出；
  4) 每个动作之后 check_dialog(pid)，有弹窗先关，再重新锚定；
  5) 提交代码后 K3 可能自动上滚 16px，此时**重新读 top** 再定位同行用量列；
  6) 绝不点「保存」（x≈1047,y≈76）。

用法：
  python _k3_fill.py probe                     # 打印窗口/当前焦点/是否有弹窗
  python _k3_fill.py calib [Y]                 # ★动态横扫标定列坐标（写 _k3_calib.json）
  python _k3_fill.py read LO HI                # 读 LO..HI 行的用量（需先滚到对应位置）
  python _k3_fill.py fill LO HI                # 录入 LO..HI 行
  python _k3_fill.py row <code> <qty> [Y]      # 在指定 y 录一行
  python _k3_fill.py clear <Y>                 # 清空该 y 所在行的代码与用量
  python _k3_fill.py scroll up|down <N>        # 用竖滚动条箭头滚 N 次
  python _k3_fill.py hscroll left|right <N>    # 用底部横向滚动条滚 N 次
"""
import sys, time, ctypes, json, os, io
from ctypes import wintypes as wt
import win32api, win32con, win32gui, win32process
from pywinauto import Application

TITLE_KW = "BOM单 - 新增"

# ---- 视口内列坐标（默认值；实际以 _k3_calib.json 为准）----
X_SEQ   = 992
X_CODE  = 1090    # 商品代码 click x（1920 宽、50/50 分屏、横拨 1 次）
X_QTY   = 1400    # 用量     click x ★ 别填 1742~1800（那是「备注」）
ROW_H   = 16

# ★ 控件左边界（用来区分 aid 相同但列不同的格）
CODE_LEFT = 1043  # 商品代码控件 rect.left
QTY_LEFT  = 1392  # 用量控件 rect.left（损耗率 1451、备注 1729，靠这个区分）
LEFT_TOL  = 14    # left 匹配容差

GRID_TOP    = 500
GRID_BOTTOM = 900
SCROLL_UP_X,   SCROLL_UP_Y   = 1908, 468
SCROLL_DOWN_X, SCROLL_DOWN_Y = 1908, 891   # ★ 竖条 rect=(1900,458,17,443)，907 在条外
HSCROLL_LEFT_X,  HSCROLL_Y   = 987, 908   # ★ 实测 (977,900,924,17) 内，旧 975/917 全落空
HSCROLL_RIGHT_X              = 1891

# 列标题带空白处：点这里可退出单元格编辑态而不改变任何数据（★ 第 11 轮）
HDR_NEUTRAL_X, HDR_NEUTRAL_Y = 1890, 476

SAVE_BTN = (1047, 76)   # 严禁点击

# set_one 实际写入用量的 y（处理 ENTER 后焦点跳首行），供 read_one 对齐
LAST_QTY_Y = None

CFG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_k3_calib.json")
if os.path.exists(CFG):
    try:
        _c = json.load(open(CFG, encoding="utf-8"))
        X_CODE = _c.get("X_CODE", X_CODE)
        X_QTY  = _c.get("X_QTY",  X_QTY)
        ROW_H  = _c.get("ROW_H",  ROW_H)
        GRID_TOP    = _c.get("GRID_TOP", GRID_TOP)
        GRID_BOTTOM = _c.get("GRID_BOTTOM", GRID_BOTTOM)
        SCROLL_UP_X, SCROLL_UP_Y = _c.get("SCROLL_UP", (SCROLL_UP_X, SCROLL_UP_Y))
        SCROLL_DOWN_X, SCROLL_DOWN_Y = _c.get("SCROLL_DOWN", (SCROLL_DOWN_X, SCROLL_DOWN_Y))
        HSCROLL_LEFT_X, HSCROLL_Y = _c.get("HSCROLL_LEFT", (HSCROLL_LEFT_X, HSCROLL_Y))
        HSCROLL_RIGHT_X = _c.get("HSCROLL_RIGHT_X", HSCROLL_RIGHT_X)
        CODE_LEFT = _c.get("CODE_LEFT", CODE_LEFT)
        QTY_LEFT  = _c.get("QTY_LEFT",  QTY_LEFT)
    except Exception as _e:
        print("calib load failed:", _e)


# ---------------- 业务数据（外部文件加载，仓库内不放真实物料） ----------------
# 数据来源优先级：环境变量 K3_DATA（路径） > 同目录 data.json > 同目录 data.csv > 空
# data.json 格式：[["物料代码", "用量"], ...]   （用量可以是字符串或数字）
# data.csv  格式：每行 `代码,用量`
import os as _os
import re
import json as _json

DATA = []
TOTAL_EXPECT = 0.0


def load_data(path=None):
    """载入待录入数据。返回 [(code, qty_str), ...]，同时刷新 TOTAL_EXPECT。"""
    global DATA, TOTAL_EXPECT
    here = _os.path.dirname(_os.path.abspath(__file__))
    cand = [path] if path else []
    if _os.environ.get("K3_DATA"):
        cand.append(_os.environ["K3_DATA"])
    cand += [_os.path.join(here, "data.json"), _os.path.join(here, "data.csv")]

    for c in cand:
        if not c or not _os.path.isfile(c):
            continue
        try:
            if c.lower().endswith(".json"):
                raw = _json.load(io.open(c, encoding="utf-8-sig"))
                rows = [(str(r[0]).strip(), str(r[1]).strip()) for r in raw if r and str(r[0]).strip()]
            else:
                rows = []
                for line in io.open(c, encoding="utf-8-sig"):
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    parts = [x.strip() for x in re.split(r"[,	;]", line)]
                    if len(parts) >= 2 and parts[0]:
                        rows.append((parts[0], parts[1]))
        except Exception as e:
            print("load_data failed on %s: %s" % (c, e))
            continue
        if rows:
            DATA = rows
            TOTAL_EXPECT = round(sum(float(q) for _, q in DATA), 4)
            print("loaded %d rows from %s, total=%.4f" % (len(DATA), c, TOTAL_EXPECT))
            return DATA
    print("WARNING: 未找到数据文件（data.json / data.csv 或 K3_DATA 环境变量），DATA 为空")
    return DATA


load_data()


# ---------------- 基础 ----------------
def find_window(kw=TITLE_KW):
    res = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and kw in win32gui.GetWindowText(h):
                res.append((h, win32process.GetWindowThreadProcessId(h)[1],
                            win32gui.GetWindowText(h), win32gui.GetWindowRect(h)))
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    res.sort(key=lambda t: -win32gui.GetWindowRect(t[0])[2])
    return res


TARGET_HWND = None

# 左侧宿主窗口（Agent / RPA 宿主）的标题关键字，可用环境变量 K3_LEFT_HINTS=foo,bar 覆盖。
# 留空 = 自动挑选：优先当前前台的那个 Chromium 系窗口，否则取最宽的可见 Chromium 窗口。
LEFT_TITLE_HINTS = ()

# ---- 前台保障（★ 第 3 轮全量自检 4/10 的根因）----
# 桌面上存在最大化的第三方浏览器 / 小组件窗口（同为 Chromium 类），会盖住整屏；
# Aero Snap 又把它带到前台，物理点击/键盘全部落空。
_u32 = ctypes.windll.user32
_u32.GetForegroundWindow.restype = wt.HWND
try:
    _u32.AttachThreadInput.argtypes = [wt.DWORD, wt.DWORD, wt.BOOL]
    _u32.AttachThreadInput.restype = wt.BOOL
    _u32.BringWindowToTop.argtypes = [wt.HWND]
    _u32.SetForegroundWindow.argtypes = [wt.HWND]
    _u32.ShowWindow.argtypes = [wt.HWND, ctypes.c_int]
    _u32.IsZoomed.argtypes = [wt.HWND]
    _u32.IsZoomed.restype = wt.BOOL
    _u32.SwitchToThisWindow.argtypes = [wt.HWND, wt.BOOL]
except Exception:
    pass
_k32 = ctypes.windll.kernel32
_k32.GetCurrentThreadId.restype = wt.DWORD


def force_foreground(hwnd, tries=3):
    """强制把 K3 置前：Windows 前台锁会让普通 SetForegroundWindow 静默失败。"""
    for _ in range(tries):
        fg = _u32.GetForegroundWindow()
        if fg == hwnd:
            return True
        cur = _k32.GetCurrentThreadId()
        fgt = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
        tgt = win32process.GetWindowThreadProcessId(hwnd)[0]
        attached = []
        for t in (fgt, tgt):
            if t and t != cur and _u32.AttachThreadInput(cur, t, True):
                attached.append(t)
        try:
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            _u32.ShowWindow(hwnd, win32con.SW_SHOW)
            _u32.BringWindowToTop(hwnd)
            _u32.SetForegroundWindow(hwnd)
            try:
                _u32.SwitchToThisWindow(hwnd, True)
            except Exception:
                pass
        finally:
            for t in attached:
                _u32.AttachThreadInput(cur, t, False)
        time.sleep(0.35)
    return _u32.GetForegroundWindow() == hwnd


def _ensure_k3_front():
    """每次物理点击/按键前保证 K3 在前台（弹窗 #32770 属同进程，无需干预）。"""
    global TARGET_HWND
    if not TARGET_HWND:
        return
    fg = _u32.GetForegroundWindow()
    if fg == TARGET_HWND:
        return
    try:
        if win32gui.GetClassName(fg) == "#32770":
            return
    except Exception:
        pass
    force_foreground(TARGET_HWND)


def connect(hwnd):
    global TARGET_HWND
    TARGET_HWND = hwnd
    app = Application(backend="uia").connect(handle=hwnd)
    return app.window(handle=hwnd)


def work_area():
    """读取系统工作区（扣除任务栏），不写死分辨率。"""
    r = wt.RECT()
    _u32.SystemParametersInfoW(0x0030, 0, ctypes.byref(r), 0)  # SPI_GETWORKAREA
    return r.left, r.top, (r.right - r.left), (r.bottom - r.top)


def _left_title_hints():
    """左侧宿主窗口的标题关键字。默认 Chromium 系；可用环境变量 K3_LEFT_HINTS 覆盖（逗号分隔）。"""
    env = os.environ.get("K3_LEFT_HINTS", "").strip()
    if env:
        return tuple(s.strip() for s in env.split(",") if s.strip())
    return LEFT_TITLE_HINTS


def ensure_layout(pid=None, verbose=True):
    """确定性把桌面摆成「宿主窗口左 / K3 右」50/50，最小化遮挡的最大化浏览器，并置前 K3。

    替代不可靠的 Aero Snap（Win+方向键会把第三方浏览器带到前台）。
    Windows API SetWindowPos 本身就是「windows 自带功能」。
    """
    r = find_window()
    if not r:
        return False, "no K3 window"
    hwnd, pid = r[0][0], r[0][1]
    left_win, intruders = [], []
    hints = _left_title_hints()

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return
            c = win32gui.GetClassName(h)
            t = win32gui.GetWindowText(h)
            if c in ("Chrome_WidgetWin_1", "Chrome_WidgetWin_0", "MozillaWindowClass"):
                r = win32gui.GetWindowRect(h)
                w = r[2] - r[0]
                if hints and any(k in t for k in hints):
                    left_win.append((h, w, h == _u32.GetForegroundWindow()))
                elif not hints and w >= 600:
                    left_win.append((h, w, h == _u32.GetForegroundWindow()))
                elif _u32.IsZoomed(h):
                    intruders.append((h, t[:40]))
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    # 无标题关键字时：前台优先，其次最宽
    if left_win and not hints:
        left_win.sort(key=lambda x: (not x[2], -x[1]))
    if left_win:
        left_win = [left_win[0][0]]
    for h, t in intruders:
        win32gui.ShowWindow(h, win32con.SW_MINIMIZE)
        if verbose:
            print("minimize intruder:", hex(h), t, flush=True)
    time.sleep(0.5)

    wx, wy, ww, wh = work_area()
    half = ww // 2

    def tile(h, x, y, w, hh):
        if _u32.IsZoomed(h):
            win32gui.ShowWindow(h, win32con.SW_RESTORE)
            time.sleep(0.4)
        win32gui.SetWindowPos(h, win32con.HWND_TOP, x, y, w, hh,
                              win32con.SWP_SHOWWINDOW)
        time.sleep(0.3)

    if left_win:
        tile(left_win[0], wx, wy, half, wh)
    tile(hwnd, wx + half, wy, ww - half, wh)
    force_foreground(hwnd)
    global TARGET_HWND
    TARGET_HWND = hwnd
    ok = _u32.GetForegroundWindow() == hwnd
    info = "left=%s K3=%s FG=%s" % (
        win32gui.GetWindowRect(left_win[0]) if left_win else None,
        win32gui.GetWindowRect(hwnd), hex(_u32.GetForegroundWindow()))
    return ok, info


def click(x, y, hold=0.06):
    _ensure_k3_front()
    win32api.SetCursorPos((int(x), int(y)))
    time.sleep(0.04)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(hold)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def wheel(notches, x=1400, y=700):
    """正数=向上，负数=向下。仅在前置 click 拿到表格焦点后有效。"""
    _ensure_k3_front()
    win32api.SetCursorPos((x, y))
    time.sleep(0.05)
    for _ in range(abs(notches)):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0,
                             120 if notches > 0 else -120, 0)
        time.sleep(0.06)


def dlg_handles(pid):
    out = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "#32770" \
                    and win32process.GetWindowThreadProcessId(h)[1] == pid:
                out.append(h)
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return out


def dlg_text(h):
    parts = []

    def cb(c, _):
        try:
            t = win32gui.GetWindowText(c)
            if t:
                parts.append(t)
        except Exception:
            pass

    win32gui.EnumChildWindows(h, cb, None)
    return " | ".join(parts)


def dialog_buttons(h):
    """返回弹窗内所有 Button 的 (文字, rect)。"""
    out = []

    def cb(c, _):
        try:
            if win32gui.GetClassName(c) == "Button":
                t = win32gui.GetWindowText(c)
                if t:
                    out.append((t, win32gui.GetWindowRect(c)))
        except Exception:
            pass

    win32gui.EnumChildWindows(h, cb, None)
    return out


# 关闭弹窗的按钮优先级：★「否」永远排第一（绝不触发保存单据），「是(&Y)」排最后
CLOSE_BTN_PRIORITY = ("否", "取消", "确定", "是")


def close_dialog(h):
    """按优先级点掉一个弹窗；返回被点按钮文字。"""
    btns = dialog_buttons(h)
    for kw in CLOSE_BTN_PRIORITY:
        for t, r in btns:
            if kw in t:
                click((r[0] + r[2]) // 2, (r[1] + r[3]) // 2)
                time.sleep(0.4)
                return t
    # 兜底：右下角
    r = win32gui.GetWindowRect(h)
    click(r[2] - 90, r[3] - 35)
    time.sleep(0.4)
    return "fallback@%s" % (r,)


def check_dialog(pid, close=True):
    """返回弹窗文字；close=True 时按文字优先点「否/取消/确定」关掉。"""
    hs = dlg_handles(pid)
    if not hs:
        return None
    txts = [dlg_text(h) for h in hs]
    if close:
        for h in hs:
            close_dialog(h)
    return " || ".join(txts)


def _esc():
    """全局发一次 ESC，撤销当前格的行内编辑态（不需要 dlg）。"""
    _ensure_k3_front()
    win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.2)


def clean_dialogs(pid, tries=4, esc=True):
    """反复关弹窗，直到没有；返回 True 表示已清干净。

    ★ 防重弹（m02439「怎么一直弹窗」的根因）：
      K3 的商品代码格逐格校验，非法代码留在编辑态时，点「否」只是不接受该值，
      编辑框文本还在 → K3 立刻再校验 → 同一个框再次弹出，形成「点否→重弹」死循环。
      因此一旦检测到**连续两次同一个弹窗文字**，必须先 ESC 撤掉该格编辑，
      校验链才会断。给所有调用点（set_one / ensure_view / reset_blank / 自检）统一兜底。
    """
    last = None
    for _ in range(tries):
        txt = check_dialog(pid)
        if txt is None:
            return True
        time.sleep(0.35)
        if esc and txt == last:
            _esc()
            if not dlg_handles(pid):
                return True
        last = txt
    if esc and dlg_handles(pid):
        for _ in range(3):
            _esc()
            check_dialog(pid)
            time.sleep(0.25)
            if not dlg_handles(pid):
                return True
    return not bool(dlg_handles(pid))


def focused(dlg):
    try:
        for e in dlg.descendants():
            try:
                if e.has_keyboard_focus():
                    r = e.rectangle()
                    return (e.automation_id(), e.window_text(),
                            e.class_name(), r.top, r.left)
            except Exception:
                continue
    except Exception:
        pass
    return None


def is_cell(f, aid, left=None, tol=None):
    """★ aid 相同不代表同一列：可选用 rect.left 精确区分。"""
    if not (f and str(f[0]) == str(aid)):
        return False
    if left is None:
        return True
    return abs(f[4] - left) <= (LEFT_TOL if tol is None else tol)


def is_code(f):
    return is_cell(f, 200, CODE_LEFT)


def is_qty(f):
    return is_cell(f, 201, QTY_LEFT)


def anchor_code(dlg, pid, y):
    """点 y 处的商品代码格，返回该行真实 focused（已处理弹窗）。

    ★ 每次点击前先 _esc()：与 detect_columns/set_one 同源——行内编辑态
      会吞掉第一次点击、把焦点留在旧格，是「anchor 命中框架/纹丝不动」的根因。
    """
    for _ in range(3):
        _esc()
        click(X_CODE, y)
        time.sleep(0.45)
        d = check_dialog(pid)
        f = focused(dlg)
        if is_code(f):
            return f
        if d and not is_code(f):
            continue
        time.sleep(0.3)
    return focused(dlg)


def submit(dlg, keys, wait=0.5):
    dlg.type_keys(keys)
    time.sleep(wait)


def set_clipboard(text):
    """把单个字符串放进剪贴板；用于数值格 ^v 粘贴，绕开 type_keys 对小数点的解析。

    ★ 第 9 轮根因：pywinauto type_keys 把主键盘 '.' (VK_OEM_PERIOD) 送进
      SPR32A70Float 数值格会被吞：'0.1182' → '01182' → K3 千分位显示 '1,182'。
      文本列 aid200 的 code 不受影响（'1.99.001' 一直正常）。
      剪贴板粘贴让控件收到完整文本而非逐键消息，小数点不再丢。
    单行单格粘贴，不是多行批量粘贴（后者会破坏 K3 自动带出）。
    """
    try:
        import win32clipboard
    except Exception:
        return False
    for _ in range(6):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, str(text))
            win32clipboard.CloseClipboard()
            return True
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            time.sleep(0.15)
    return False


def safe_keys(dlg, pid, keys, retry=2, wait=0.0):
    """type_keys 在模态弹窗期间会抛 ElementNotEnabled；先清框再重试。"""
    _ensure_k3_front()
    for _ in range(retry + 1):
        try:
            dlg.type_keys(keys)
            if wait:
                time.sleep(wait)
            return True
        except Exception:
            clean_dialogs(pid)
            time.sleep(0.4)
    return False


# ---------------- 动态标定 ----------------
def _is_float_col(txt):
    """aid=201 的格都是数值列；用 rect.left 顺序区分，不依赖小数位数。

    实测（归零后右拨 1 次）：用量 left=1654 值 '1.0000' / 损耗率 left=1756 值 '0.00'；
    归零后右拨 2 次：用量 left=1274 / 损耗率 left=1376。
    ⚠️ 首行是空行时用量值可能显示成 '0'（无小数点），所以**不能靠 4 位小数判定**，
       必须按 aid201 里 rect.left 最小者取用量列。
    """
    t = (txt or "").strip()
    if t == "":
        return False
    try:
        float(t.replace(",", ""))
        return True
    except ValueError:
        return False


def detect_columns(dlg, pid, y=None, home=True):
    """★ 固定视口探测（第 5 轮重写）：归零 + 每次点击前 ESC + 固定坐标单点验证。

    ★ 弃用「24px 步长连续横扫」的决定性理由（m02600/m02607）：
      Spread 控件每被点中一个靠右格子就进入行内编辑态并**自动水平滚动**，
      所有列 rect.left 随之平移约 380px；「边扫边测」测到的是不断漂移的中间态，
      算出 QTY_LEFT=1281 或 1661 全凭起始位置，验证点永远对不上——自毁式探测。
      _k3_probe_fixed.py 证明：**每个 click 前加 _esc()（只选中不进编辑态）后，
      归零态 step0/1/2 三次结果完全一致，视口不再漂移**。

    ★ 归零态稳定列分布（_k3_probe_fixed.py，y=492/500，K3 50/50 分屏）：
        商品代码 aid200 left=1043  (x=1043/1092)
        用量     aid201 left=1281  (x=1281/1654, 值 '0')
        文本     aid200 left=1673  (x=1701/1722/1763)
      y=508 是行边界会弹「金蝶提示」，探测 y 必须 <= 504。
      ⚠️ 归零态视口里只有一列 aid201（用量），损耗率在视口外，因此不会误点损耗率。

    ★ 与 GRID_TOP 解耦：cfg 的 GRID_TOP 固定写 492（实测首行 top），
      绝不写成 y-8，否则下次默认探测 y 会漂移到列标题区 484。

    绝不修改任何模块全局、绝不写文件；只把横向滑条归零使视口确定。
    返回 cfg dict 或 None（固定坐标验证失败）。
    """
    hwnd, pid2, title, rect = find_window()[0]
    py = y or 496                       # 首行格中心，避开 508 弹窗边界
    if home:
        hscroll_home(pid, steps=12, dlg=dlg)  # 视口确定：归零态（ensure_view 循环里由调用方统一拨）
    clean_dialogs(pid, tries=4)

    # ESC + 横扫（★ 每个 click 前 _esc() 只选中不进编辑态 → 视口不漂移）
    # _k3_probe_xs.py 实测 y=492/496/500 三个 y 横扫结果完全一致，证明稳定。
    # 归零态列分布：商品代码 aid200/1043 | 用量 aid201/1281 '0'
    #              | 损耗率 aid201/1383 '0.00' | 文本 aid200/1532, 1673
    # ⇒ 用量 = aid201 中 rect.left 最小者；绝不硬编码 x（列宽会变，
    #   m02630 曾误把 x=1467 当用量，实际那是损耗率 left=1383）。
    # ★ 第 12 轮（m03763 实测）：横扫过程中「点靠右的格子」会触发 K3 自动
    #   横向滚动，视口整体平移 ⇒ 同一个 x 在不同时刻命中不同列的 left
    #   （实测 x=1360 有时命中 left=1349、有时命中 left=1729）。
    #   过滤判据：命中可信 ⟺ 列左边界不晚于点击 x，且 x 未超出该列宽度太多
    #   （`left <= x <= left + 200`）。不可信的命中直接丢弃，不参与分组。
    # ★ 第 12 轮（m03797 用户要求「别扫了，效率高点」）：
    #   ★ 关键：横扫**只在启动时做一次**（cmd_fill 开头 calib 一次），
    #     36 行录入过程里绝不再扫 —— 行内的 recover 走 quick_fix_view（≤6 次点击）。
    #   步长从 24 加大到 48（18 点而非 35 点），单次标定从 ~10s 降到 ~5s。
    seen = {}
    step_x = 48
    x = rect[0] + 40
    while x <= rect[2] - 30:
        _esc()
        click(x, py)
        time.sleep(0.1)
        f = focused(dlg)
        if f and str(f[0]) in ("200", "201"):
            left = f[4]
            if left <= x <= left + 200:          # 几何自洽才算真命中
                g = seen.setdefault((str(f[0]), left), {"xs": [], "val": f[1]})
                g["xs"].append(x)
                if f[1]:
                    g["val"] = f[1]
        x += step_x
    clean_dialogs(pid, tries=4)
    # 稀疏探测每列可能只命中 1 个点，阈值降到 1（横扫时代的 >=2 是为过滤假组）
    seen = {k: v for k, v in seen.items() if len(v["xs"]) >= 1}
    codes = sorted([(k[1], v) for k, v in seen.items() if k[0] == "200"])
    # ★ 第 12 轮（m03756）：绝不能用 _is_float_col 过滤 aid201！
    #   空白单据里「用量」列显示的是**空字符串 ''**（不是 '0'），
    #   一旦按「能转 float」筛，用量列被丢掉，只剩「损耗率 '0.00'」，
    #   于是 QTY_LEFT=1451 指向损耗率、数据全写错列（截图+OCR 证实
    #   用量标题在 x≈1394、损耗率在 x≈1467）。改为**只按 rect.left 最小**取用量列。
    #   实测同一视口下 aid201 的 left 集合恒为 {1349(用量), 1451(损耗率)}。
    qtys = sorted([(k[1], v) for k, v in seen.items() if k[0] == "201"])
    if not codes or not qtys:
        return None
    # ★ 第 10 轮（截图 + UIA 横扫 + _k3_probe_ensureview.py 逐步 instrument 实证）：
    #   横扫命中 x 的「首尾中点」会随滑条起始位置 / ESC 时序左右漂移 ——
    #   实测 step0=1324 / step1=1180 / step2=1324 / step3=1180 交替，
    #   1180 落在文本列 aid200/1673 上，验证点击被吞 → ensure_view 误判 False，
    #   连锁把 single_row / multi_row / columns_visible 全打挂（8/10 的最后一环）。
    #   控件真实 rect.left 是 UIA 属性、不受横扫起点影响；点 left+8 必落在该列格内
    #   （_k3_shot_steps.py 归零态实测：用量 left=1281，x=1280/1320/1360 全命中 aid201）。
    #   ⇒ X_CODE / X_QTY 一律由 CODE_LEFT / QTY_LEFT 推导，绝不再用 xs 中点。
    return {
        "X_CODE": codes[0][0] + 8,
        "CODE_LEFT": codes[0][0],
        "X_QTY": qtys[0][0] + 8,
        "QTY_LEFT": qtys[0][0],
        "Y_PROBE": py, "GRID_TOP": 492,   # 固定首行 top，绝不写 y-8 漂移
        "groups200": [c[0] for c in codes],
        "groups201": [(q[0], q[1]["val"]) for q in qtys],
        "hwnd": hwnd, "pid": pid, "rect": list(rect),
    }


def apply_calib(cfg, write=True):
    """把探测结果提交到模块全局；write=True 时落盘 _k3_calib.json。"""
    global X_CODE, X_QTY, CODE_LEFT, QTY_LEFT, GRID_TOP
    X_CODE = cfg["X_CODE"]
    X_QTY = cfg["X_QTY"]
    CODE_LEFT = cfg["CODE_LEFT"]
    QTY_LEFT = cfg["QTY_LEFT"]
    GRID_TOP = cfg["GRID_TOP"]
    if write:
        out = {
            "X_CODE": X_CODE, "X_QTY": X_QTY, "ROW_H": ROW_H,
            "GRID_TOP": GRID_TOP, "GRID_BOTTOM": GRID_BOTTOM,
            "SCROLL_UP": [SCROLL_UP_X, SCROLL_UP_Y],
            "SCROLL_DOWN": [SCROLL_DOWN_X, SCROLL_DOWN_Y],
            "HSCROLL_LEFT": [HSCROLL_LEFT_X, HSCROLL_Y], "HSCROLL_Y": HSCROLL_Y,
            "HSCROLL_RIGHT_X": HSCROLL_RIGHT_X,
            "CODE_LEFT": CODE_LEFT, "QTY_LEFT": QTY_LEFT,
            "HSCROLL_RIGHT_STEPS": 1,
            "window": [cfg["hwnd"], cfg["pid"], cfg["rect"]],
            "note": "detect_columns: aid200/201 非列唯一，按 rect.left 取最小者；归零后右拨1次",
        }
        json.dump(out, open(CFG, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)


def auto_calibrate(dlg, pid, y=None, write=False, home=True):
    """兼容旧名：探测列位置，返回 (ok, info, cfg)。

    ★ write=False 时不改任何全局、不写文件，彻底消除「ensure_view 每步
      auto_calibrate(write=False) 把 QTY_LEFT 污染成 1274」的根因。
      调用方验证命中后再显式 apply_calib(cfg)。
    """
    cfg = detect_columns(dlg, pid, y, home=home)
    if cfg is None:
        return False, "用量列不在视口，先归零再右拨 1 次", None
    info = "X_CODE=%d CODE_LEFT=%d X_QTY=%d QTY_LEFT=%d groups200=%r groups201=%r" % (
        cfg["X_CODE"], cfg["CODE_LEFT"], cfg["X_QTY"], cfg["QTY_LEFT"],
        cfg["groups200"], cfg["groups201"])
    if write:
        apply_calib(cfg, write=True)
    return True, info, cfg


# ★ 第 11 轮（UIA 实证，m03472）：底部横向 ScrollBar 的真实 rect 是
#   (977, 900, 924, 17) —— 可点中心 y=908。旧常量 HSCROLL_Y=917 落在
#   滚动条下边框之外，点击全部落空 ⇒ 视口永远卡在 HOME，用量列永远在视口外，
#   calib 必然 FAIL。而列宽可被用户手动拉宽（商品名称/规格型号变宽），
#   箭头位置也随之变化 ⇒ 必须每次从 UIA 读取真实几何，绝不写死。
def vsb_geom(dlg=None):
    """返回 (up_x, up_y, down_x, down_y)：右侧竖向 ScrollBar 的上下箭头可点坐标。

    ★ 第 12 轮（m03680）：UIA 实测竖向 ScrollBar `rect=(1900,458,1917,901)`，
      旧常量 SCROLL_DOWN=(1910,907) 的 y=907 **在滚动条下端 901 之外**
      ⇒ 下滚箭头从来没被点中，26 行以后的下滚永远是「点了没动」。
    """
    fb = (1908, 470, 1908, 893)
    if dlg is None:
        return fb
    try:
        best = None
        for c in dlg.descendants():
            try:
                if c.class_name() != "ScrollBar":
                    continue
                r = c.rectangle()
                w, h = r.right - r.left, r.bottom - r.top
                if w <= 30 and h > 100:          # 竖向：窄而高
                    if best is None or h > (best[3] - best[1]):
                        best = (r.left, r.top, r.right, r.bottom)
            except Exception:
                continue
        if best:
            cx = (best[0] + best[2]) // 2
            return (cx, best[1] + 10, cx, best[3] - 10)
    except Exception:
        pass
    return fb


def vscroll(pid, direction, n=1, dlg=None, wait=0.5):
    """竖向下/上滚 n 格（点滚动条箭头，每格约 1 行）。"""
    ux, uy, dx, dy = vsb_geom(dlg)
    x, y = (ux, uy) if direction == "up" else (dx, dy)
    for _ in range(n):
        click(x, y); time.sleep(wait); clean_dialogs(pid, tries=2)


def hsb_geom(dlg=None):
    """返回 (left_x, right_x, y)：底部横向 ScrollBar 的左右箭头可点坐标。

    ★ 第 11 轮修正（m03578）：K3 的滚动条 control_type 就是 "ScrollBar"，
      旧实现按 control_type="Pane" 过滤 → 一个都取不到、永远走 fallback
      （fallback 又是错值 975/917）。改成只按 class_name 匹配。
    """
    fb = (987, 1891, 908)
    if dlg is None:
        return fb
    try:
        best = None
        for c in dlg.descendants():
            try:
                if c.class_name() != "ScrollBar":
                    continue
                r = c.rectangle()
                w, h = r.right - r.left, r.bottom - r.top
                if h <= 30 and w > 200:          # 横向：矮而宽
                    if best is None or w > (best[2] - best[0]):
                        best = (r.left, r.top, r.right, r.bottom)
            except Exception:
                continue
        if best:
            return (best[0] + 10, best[2] - 10, (best[1] + best[3]) // 2)
    except Exception:
        pass
    return fb


def hscroll_home(pid, steps=12, dlg=None):
    """把底部横向滑条拨回最左（归零）。K3 点一次位移很大，12 次足够到头。"""
    lx, rx, y = hsb_geom(dlg)
    for _ in range(steps):
        click(lx, y)
        time.sleep(0.2)
        clean_dialogs(pid, tries=1)
    time.sleep(0.3)


def hscroll_right(pid, n=1, dlg=None, wait=0.45):
    """向右拨 n 次（每次按 UIA 实测的右箭头坐标）。"""
    lx, rx, y = hsb_geom(dlg)
    for _ in range(n):
        click(rx, y)
        time.sleep(wait)
        clean_dialogs(pid)


def quick_fix_view(dlg, pid, y):
    """轻量视口恢复：不横扫、不全量标定，只按当前 X_QTY 试几个 x 偏移。

    ★ 第 12 轮（m03786 用户实测反馈「没必要每行都扫」）：
      detect_columns 的一次横扫要点击 ~35 次（每次 0.12s + clean_dialogs），
      单行就要 8~15 秒；36 行要 6 分钟以上，而且横扫本身会触发自动横滚、
      把视口拨乱。实测横向视口一旦标定好，整份单据录入期间**不会自己变**
      （只有 reset_blank() 新增→否 会重置）。所以把「命中失败」的恢复
      降级为：按 QTY_LEFT 在当前视口附近的几个候选 x 试一下（每个 x 只点
      一次，共 4 次），全不中才回落到全量 ensure_view。
    """
    cands = []
    for dx in (0, -8, +42, -48, +90, -90):
        x = X_QTY + dx
        if x not in cands:
            cands.append(x)
    for x in cands:
        _esc()
        click(x, y); time.sleep(0.3)
        clean_dialogs(pid, tries=2)
        f = focused(dlg)
        if f and str(f[0]) == "201":
            # 命中数值列，且 left 与我们记录的 QTY_LEFT 一致才算数
            if abs(f[4] - QTY_LEFT) <= 200:
                return True, f, x
    return False, None, None


def ensure_view(dlg, pid, y=None, write=True):
    """★ SKILL 步骤 2：归零 → 逐步右拨 0..3 次 → 每步 ESC 固定坐标探测 + 单点复核。

    ★ 第 9 轮（m02810）恢复循环的根因：NB.reset() 后横向滑条起始位置不确定，
      归零态**有时**只有 aid200 铺满、用量 aid201 在视口外（第 2 轮诊断原话），
      有时用量已在视口内。第 5 轮把「归零右拨」循环删掉只探一次 → 时好时坏，
      columns_visible / calib_hit 开局随机挂。
    ★ 关键：detect_columns(home=False) 在本循环里**不再自己归零**，由 ensure_view
      统一控制滑条；每步横扫仍「每点前 _esc()」防漂移（第 5 轮决定性结论）。
    ★ 探测纯只读（write=False 不改全局）；只有单点复核 is_cell(f,201,QTY_LEFT)
      命中才 apply_calib 一次提交，避免第 2 轮的全局污染。
    """
    # ★ 第 11 轮（m03540）：进循环前必须先退出单元格编辑态。
    #   set_one 里若已经点过格子并留下 caret，滚动条箭头的点击会被 K3 表格吞掉，
    #   hscroll_home/right 全部落空 → 4 步探测全 FAIL（单行用例实测）。
    _esc(); _esc()
    click(HDR_NEUTRAL_X, HDR_NEUTRAL_Y)      # 点列标题空白处，彻底退编辑态
    time.sleep(0.2)
    clean_dialogs(pid)

    hscroll_home(pid, steps=12, dlg=dlg)   # ★ 先归零一次，固定起点
    for step in range(0, 4):             # step0=归零态，之后每步右拨 1 次
        ok, info, cfg = auto_calibrate(dlg, pid, y=y, write=False, home=False)
        if ok and cfg:
            _esc()
            click(cfg["X_QTY"], cfg["Y_PROBE"])
            time.sleep(0.4)
            clean_dialogs(pid)
            f = focused(dlg)
            if is_cell(f, 201, cfg["QTY_LEFT"]):
                apply_calib(cfg, write=write)
                return True, "right_step=%d %s hit=%r" % (step, info, f)
        # 本步没命中 → 右拨 1 次再探（★ 第 11 轮：必须按 UIA 实测几何点）
        _esc()
        hscroll_right(pid, 1, dlg=dlg)
    return False, "归零后右拨 3 次仍无法把用量列稳定拨进视口"


def clear_cell(dlg, pid, x, y, want_aid=200):
    """清空 (x,y) 处单元格，返回 (ok, focused)。

    ★ 第 11 轮（m03596 t_clear FAIL 根因）：旧 t_clear 直接 click 而不先 _esc()，
      行内编辑态会吞掉第一次点击、焦点留在上一行/上一列，^a+DELETE 删错格，
      读回 1.0000/2.0000/0.1182 依然存在。这里统一「ESC → 点击 → 校验命中
      → ^a → DELETE → ENTER → 清弹窗」，最多 2 轮。
    """
    f = None
    for _ in range(2):
        _esc()
        click(x, y); time.sleep(0.4)
        clean_dialogs(pid)
        f = focused(dlg)
        if f and str(f[0]) == str(want_aid):
            break
    if not (f and str(f[0]) == str(want_aid)):
        return False, f
    safe_keys(dlg, pid, "^a", wait=0.12)
    safe_keys(dlg, pid, "{DELETE}", wait=0.12)
    safe_keys(dlg, pid, "{ENTER}", wait=0.9)
    clean_dialogs(pid)
    _esc()
    return True, f


def bring_qty_into_view(dlg, pid, max_steps=8, write=True):
    """兼容旧调用名：等价于 ensure_view。"""
    return ensure_view(dlg, pid, write=write)


# ---------------- 业务 ----------------
def set_one(dlg, pid, code, qty, y):
    """在 y 行录入 code+qty，全程锚定。返回 (ok, 说明)。

    ★ 第 6 轮重写（m02642 单/多行 FAIL + 400s 超时的根因）：
      1) 点 X_QTY 前必须 _esc()：code 提交 ENTER 后仍在行内编辑态，
         第一次点击被吞、焦点弹回代码格（与 detect_columns 自毁同源）。
      2) 提交 code 后绝不信任 focused() 的 top：clean_dialogs 的 ESC
         会把焦点打回 top=492 首行，旧代码据此算 y_qty=500，把 3 行
         全写进首行，read_one 却在 516/532/548 读空行 '0'。
         改成重新点 X_CODE,y 拿该行真实 top2，再算 y_qty=top2+8。
      3) 命中不了时不再 auto_calibrate（归零+40步横扫 ≈ 20s/次，
         反复触发把整轮拖超时）；只 _esc() + 重试 + 16px 微调。
      4) 实际写入的 y 记入全局 LAST_QTY_Y，供 read_one/selftest 对齐。
    ⚠️ 所有按键走 safe_keys；判定列必须 (aid, rect.left) 双条件。
    """
    global LAST_QTY_Y
    LAST_QTY_Y = None
    _esc()
    f = anchor_code(dlg, pid, y)
    if not is_code(f):
        return False, "未命中代码格(aid200,left=%d): %r" % (CODE_LEFT, f)
    top = f[3]
    if not safe_keys(dlg, pid, "^a", wait=0.15):
        return False, "代码格 ^a 失败"
    if not safe_keys(dlg, pid, code, wait=0.25):
        return False, "代码格输入失败"
    if not safe_keys(dlg, pid, "{ENTER}", wait=1.6):
        return False, "代码格 ENTER 失败"
    if clean_dialogs(pid) is False:
        return False, "弹窗无法关闭"
    d = check_dialog(pid, close=False)
    if d:
        return False, "弹窗: %s" % d

    # 重新锚定：点回 X_CODE,y 拿该行真实 top（不用提交后的 focused）
    _esc()
    f2 = anchor_code(dlg, pid, y)
    top2 = f2[3] if is_code(f2) else top
    if top2 != top:
        print("    (row re-anchor: top %d -> %d)" % (top, top2), flush=True)

    fq = None
    y_qty = top2 + 8
    x_qty = X_QTY
    for attempt in (0, 1):
        for dy in (0, -ROW_H, ROW_H):
            y_try = top2 + 8 + dy
            _esc()
            click(x_qty, y_try); time.sleep(0.4)
            clean_dialogs(pid, tries=2)
            fq = focused(dlg)
            if is_qty(fq):
                y_qty = y_try
                break
        if is_qty(fq):
            break
        if attempt:
            break
        # ★ 第 12 轮（m03786）：先用轻量 quick_fix_view（4 次点击），
        #   全量横扫标定只在它也不成功时才做（横扫 ~35 次点击、很慢且会扰动视口）。
        okq, fq2, xq2 = quick_fix_view(dlg, pid, top2 + 8)
        if okq:
            fq = fq2
            x_qty = xq2
            y_qty = top2 + 8
            print("    (qty col quick-fix -> x=%d %r)" % (xq2, fq2), flush=True)
            break
        okv, infov = ensure_view(dlg, pid, y=top2 + 8, write=True)
        print("    (qty col lost -> ensure_view: %s %s)" % (okv, infov), flush=True)
        if not okv:
            break
        x_qty = X_QTY
        _esc()
        f2b = anchor_code(dlg, pid, y)
        if is_code(f2b):
            top2 = f2b[3]
    if not is_qty(fq):
        return False, "未命中用量格(aid201,left=%d): %r" % (QTY_LEFT, fq)

    # ★ 第 9 轮（_k3_probe_input.py 实证）：type_keys 把裸 '.' 吞掉，
    #   '0.1182' → '01182' → K3 显示 '11,821.0000' / '1,182'；
    #   用 {.} 显式转义小数点后 committed='0.1182' 完整
    #   （B {VK_DECIMAL} / C {.} / D 剪贴板^v / E raw VK_DECIMAL 四种写法均验证通过）。
    #   选 {.}：无剪贴板依赖、不抢前台、单行单格、整数 qty 不受影响。
    # ★ 第 11 轮（m03588）：仍偶发被吞（multi_row 第 4 行 '0.1182'→'1,182'），
    #   根因是上一格状态未提交干净、^a 选中范围没覆盖全值。改成「写→读回→
    #   不符就 ^a 清空重写」最多 3 轮，每轮都在点击前重新锚定该格。
    qty_keys = str(qty).replace(".", "{.}")
    want = float(str(qty))
    committed = None
    for wi in range(3):
        for _ in range(2):                       # 确保焦点在该用量格内
            _esc()
            click(x_qty, y_qty); time.sleep(0.35)
            clean_dialogs(pid)
            if is_qty(focused(dlg)):
                break
        if not safe_keys(dlg, pid, "^a", wait=0.12):
            continue
        # ★ 第 11 轮：^a 后直接打字会把旧值留在前面（'1' → '11.0000'），
        #   所以每轮都先 DELETE 清成空值，再 ^a 一次、再输入，避免拼接。
        safe_keys(dlg, pid, "{DELETE}", wait=0.12)
        safe_keys(dlg, pid, "^a", wait=0.1)
        if not safe_keys(dlg, pid, qty_keys, wait=0.28):
            continue
        if not safe_keys(dlg, pid, "{ENTER}", wait=0.9):
            continue
        clean_dialogs(pid)
        d = check_dialog(pid, close=False)
        if d:
            return False, "弹窗: %s" % d
        committed = read_one(dlg, pid, y_qty)
        LAST_QTY_Y = y_qty
        got = None
        if committed is not None:
            try:
                got = float(str(committed).replace(",", "").strip())
            except Exception:
                got = None
        if got is not None and abs(got - want) <= 1e-4:
            return True, "top=%d y_qty=%d qty=%s read=%r left=%d write#%d" % (
                top2, y_qty, qty, committed, fq[4], wi + 1)
        if wi == 2:
            break
        # 清空该格再重写：全选后 DELETE，让 K3 回到空值，避免拼接旧数字
        _esc()
        click(x_qty, y_qty); time.sleep(0.3)
        safe_keys(dlg, pid, "^a", wait=0.1)
        safe_keys(dlg, pid, "{DELETE}", wait=0.2)
        safe_keys(dlg, pid, "{ENTER}", wait=0.5)
        clean_dialogs(pid)
        print("    (qty retry #%d: 写%s 读%s)" % (wi + 1, qty, committed), flush=True)
    return False, "用量读回不符: 写%s 读%s (y=%d)" % (qty, committed, y_qty)


def read_one(dlg, pid, y=None):
    """读 y 行用量值；y=None 时读 set_one 实际写入的 LAST_QTY_Y。

    ★ 第 6 轮：点 X_QTY 前 _esc()（编辑态会吞点击），命中不了只
      ESC 重试，不再 auto_calibrate（归零+横扫会把视口拨乱且极慢）。
    """
    if y is None:
        y = LAST_QTY_Y
    if y is None:
        return None
    for i in range(3):
        _esc()
        click(X_QTY, y); time.sleep(0.4)
        clean_dialogs(pid)
        f = focused(dlg)
        if is_qty(f):
            return f[1]
        if i == 1:
            # ★ 视口被重置（新增/翻页）时自愈一次，避免读回 None 误判为写入失败
            okv, infov = ensure_view(dlg, pid, y=y, write=True)
            print("    (read lost -> ensure_view: %s %s)" % (okv, infov), flush=True)
            if not okv:
                break
            y = y or (GRID_TOP + 8)
    return None


def last_visible_y():
    """当前视口最后一个**有效代码行**的 y。

    ★ 第 7 轮（auto_scroll FAIL 根因）：旧值 886 落在表格底边框之外，
      点击命中 SPR32A70_SpreadSheet 框架 (top=886,left=978) 而非代码格。
      纵扫实测：GRID_TOP+8 .. 884 全部命中 aid200/1043，886 越界。
      取 GRID_BOTTOM 前一个整行，且用 (GRID_TOP+8) 起算的行栅格对齐。
    """
    top = GRID_TOP + 8                       # 500
    last = GRID_BOTTOM - ROW_H               # 884（诊断确认的最后一个可命中行）
    n = (last - top) // ROW_H
    return top + n * ROW_H                   # 884


def cmd_fill(lo, hi, start_y=None):
    hwnd, pid, title, rect = find_window()[0]
    print("window:", title, hex(hwnd), pid, rect, flush=True)
    dlg = connect(hwnd)
    clean_dialogs(pid)
    try:
        dlg.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    ok, info = bring_qty_into_view(dlg, pid)
    print("calib:", ok, info, flush=True)
    if not ok:
        return
    y = start_y if start_y else (GRID_TOP + 8 + ROW_H * (lo - 1))
    for n in range(lo, hi + 1):
        code, qty = DATA[n - 1]
        ok, info = set_one(dlg, pid, code, qty, y)
        print("row %2d %-16s %-8s -> %s %s" % (n, code, qty, "OK " if ok else "FAIL", info), flush=True)
        if not ok:
            print("!! aborted at row", n, flush=True); return
        f = focused(dlg)
        if f and str(f[0]) in ("200", "201"):
            y = f[3] + 8 + ROW_H
        else:
            y = y + ROW_H
        if y > GRID_BOTTOM:
            print("-- viewport bottom reached, scroll down via vsb_geom --", flush=True)
            # ★ 第 12 轮（m03703）：旧代码写死 click(SCROLL_DOWN_X, SCROLL_DOWN_Y)
            #   =(1910,907)，而实测竖向滚动条 rect=(1900,458,1917,901)，
            #   y=907 在下端 901 之外 ⇒ 从来没滚动过，26 行以后全部重复写末行。
            vscroll(pid, "down", 1, dlg=dlg, wait=0.6)
            y = last_visible_y()
            print("   after scroll: y=%d vsb=%r" % (y, vsb_geom(dlg)), flush=True)
    print("done. TOTAL should sum to", sum(float(q) for _, q in DATA[lo-1:hi]), flush=True)


# ---------------- CLI ----------------
def main():
    a = sys.argv[1:]
    if not a:
        print(__doc__); return
    cmd = a[0]
    if cmd == "probe":
        for h, pid, t, r in find_window():
            print(hex(h), pid, t, r, flush=True)
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        print("focused:", focused(dlg), flush=True)
        print("dialog:", check_dialog(pid, close=False), flush=True)
        print("calib: X_CODE=%d X_QTY=%d CODE_LEFT=%d QTY_LEFT=%d" % (
            X_CODE, X_QTY, CODE_LEFT, QTY_LEFT), flush=True)
    elif cmd == "calib":
        y = int(a[1]) if len(a) > 1 else None
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        clean_dialogs(pid)
        try: dlg.set_focus()
        except Exception: pass
        time.sleep(0.2)
        ok, info = bring_qty_into_view(dlg, pid)
        print(("OK" if ok else "FAIL"), info, flush=True)
    elif cmd == "read":
        lo, hi = int(a[1]), int(a[2])
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        clean_dialogs(pid)
        try: dlg.set_focus()
        except Exception: pass
        base = GRID_TOP + 8
        for i, n in enumerate(range(lo, hi + 1)):
            print(n, read_one(dlg, pid, base + ROW_H * i), flush=True)
    elif cmd == "fill":
        cmd_fill(int(a[1]), int(a[2]))
    elif cmd == "row":
        code, qty = a[1], a[2]
        y = int(a[3]) if len(a) > 3 else last_visible_y()
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        clean_dialogs(pid)
        try: dlg.set_focus()
        except Exception: pass
        time.sleep(0.2)
        print(set_one(dlg, pid, code, qty, y), flush=True)
    elif cmd == "clear":
        y = int(a[1])
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        clean_dialogs(pid)
        try: dlg.set_focus()
        except Exception: pass
        time.sleep(0.2)
        for x in (X_QTY, X_CODE):
            click(x, y); time.sleep(0.5); clean_dialogs(pid)
            try:
                dlg.type_keys("^a"); time.sleep(0.15)
                dlg.type_keys("{DELETE}"); time.sleep(0.15)
                dlg.type_keys("{ENTER}"); time.sleep(1.5)
            except Exception as e:
                print("clear key error:", e, flush=True)
            clean_dialogs(pid)
            print("cleared x=%d ->" % x, focused(dlg), flush=True)
    elif cmd == "scroll":
        n = int(a[2]) if len(a) > 2 else 1
        hwnd0, pid0, t0, r0 = find_window()[0]
        dlg0 = connect(hwnd0)
        vscroll(pid0, "up" if a[1] == "up" else "down", n, dlg=dlg0)
        print("scrolled", a[1], n, "vsb=", vsb_geom(dlg0), flush=True)
    elif cmd == "hscroll":
        n = int(a[2]) if len(a) > 2 else 1
        hwnd, pid, t, r = find_window()[0]
        dlg = connect(hwnd)
        if a[1] == "left":
            hscroll_home(pid, steps=n, dlg=dlg)
        else:
            hscroll_right(pid, n, dlg=dlg)
        print("hscrolled", a[1], n, "geom=", hsb_geom(dlg), flush=True)
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
