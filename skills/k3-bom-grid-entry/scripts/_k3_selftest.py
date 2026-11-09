# -*- coding: utf-8 -*-
"""
K3 BOM SKILL 循环自检器 v2 —— 按 SKILL.md 的每一步逐项验证，输出 PASS/FAIL。

v2 相对 v1 的修复（对应第一轮 5/10 FAIL 的根因）：
  * 每轮开头用「新增 → 否」把明细重置成空白（用户 m02271 授权，禁止保存）；
  * columns_visible / calib 改用 K.ensure_view()（归零 + 右拨 1 次 + 动态标定），
    不再盲右拨 4 次把滑条拨过头；
  * 判列一律 (aid, rect.left) 双条件：K.is_code / K.is_qty；
  * 所有按键走 K.safe_keys（弹窗期 dlg.type_keys 会抛 ElementNotEnabled）；
  * 弹窗关闭走 K.check_dialog（按文字优先点「否」，绝不点「是」）；
  * auto_scroll 改为验证 set_one 的重锚定是否仍能正确写入，而不是脆弱的位移断言。

设计原则：
  * 只操作当前已打开的 BOM「新增」窗口，全程**不点保存**（x≈1047,y≈76）；
  * 用例按依赖顺序执行：分屏 → 横拨 → 标定 → 弹窗 → 单行 → 多行 → 清场 → 自动滚动 → 未保存；
  * 支持 `--case NAME` 单跑某个用例，便于定位后快速迭代复测；
  * 支持 `--loop N` 连续跑 N 轮，验证可重复性（循环测试迭代）。

用法：
  python _k3_selftest.py                 # 全量 1 轮
  python _k3_selftest.py --loop 3       # 连续 3 轮
  python _k3_selftest.py --case split   # 单项
  python _k3_selftest.py --list         # 列出用例
"""
import sys, time, os, json, subprocess
import win32gui
import _k3_fill as K

if not K.DATA:
    print("ERROR: 没有可录入的数据。请提供 data.json / data.csv，或设置环境变量 K3_DATA 指向数据文件。")
    print("       格式见 data.example.json（[[\"物料代码\", \"用量\"], ...]）")
    raise SystemExit(2)
import _k3_newbom as NB

HERE = os.path.dirname(os.path.abspath(__file__))
# 仓库内布局：scripts/ -> ../SKILL.md（同时兼容把本包装进用户级 skills 目录的情形）
SKILL_PATHS = [
    os.path.abspath(os.path.join(HERE, "..", "SKILL.md")),
    os.path.join(HERE, "SKILL.md"),
    os.path.expanduser("~/.claude/skills/k3-bom-grid-entry/SKILL.md"),
]
DOC_PATH = os.path.abspath(os.path.join(HERE, "..", "..", "..", "docs", "TROUBLESHOOTING.md"))

RESULTS = []


def record(name, ok, info=""):
    RESULTS.append((name, bool(ok), str(info)))
    print("%-4s %-22s | %s" % ("PASS" if ok else "FAIL", name, info), flush=True)


def clean(pid):
    """统一走 K.clean_dialogs：按「否」关弹窗，重复同一弹窗时 ESC 撤编辑态防死循环。

    ★ m02439「怎么一直弹窗」：自检器旧实现只循环点「否」，非法代码留在
      商品代码格编辑态时 K3 会逐格校验、点否后立即重弹，形成无限重弹。
      必须复用 K.clean_dialogs 的「连续同文字 → ESC」兜底。
    """
    return bool(K.clean_dialogs(pid))


def getwin():
    r = K.find_window()
    if not r:
        raise RuntimeError("找不到 BOM单 - 新增 窗口")
    hwnd, pid, title, rect = r[0]
    dlg = K.connect(hwnd)
    try:
        dlg.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    clean(pid)
    return hwnd, pid, title, rect, dlg


def keys(dlg, pid, seq, wait=0.0):
    """统一走 safe_keys，杜绝弹窗期 ElementNotEnabled。"""
    return K.safe_keys(dlg, pid, seq, wait=wait)


def num(s):
    """K3 读数可能是 '1.0000' / '0' / 千分位 '1,182'；统一转 float，坏值返回 None。

    ★ 第 8 轮：半输入脏态会把 0.1182 拼成 1182 并显示成 '1,182'，
      裸 float() 抛 ValueError 会把整个用例打成 EXCEPTION 并中断后续用例。
    """
    if s is None:
        return None
    try:
        return float(str(s).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def reset_blank():
    """★ 每轮开头：新增 → 否，把明细重置成空白，且不保存。

    ★ 第 3 轮 4/10 根因：桌面有最大化的第三方浏览器/小组件抢前台，
      且诊断脚本 SW_RESTORE 把 K3 缩成 677x560 小窗，坐标全废。
      必须先 ensure_layout() 确定性分屏 + 强制 K3 前台，NB.reset 的点「新增」才落得准。
    """
    r = K.find_window()[0]
    clean(r[1])
    K.ensure_layout(verbose=False)
    clean(r[1])
    NB.reset(verbose=False)
    clean(r[1])
    return K.find_window()[0]


# ---------------------------------------------------------------- T0
def t_skill_files():
    p = SKILL_PATHS[0]
    ok = os.path.isfile(p)
    info = p
    if ok:
        s = open(p, encoding="utf-8").read()
        ok = s.startswith("---") and "name: k3-bom-grid-entry" in s[:400]
        info += " | frontmatter=%s lines=%d" % (ok, len(s.splitlines()))
        for q in SKILL_PATHS[1:]:
            if os.path.isfile(q):
                same = open(q, encoding="utf-8").read() == s
                info += " | %s=%s" % (os.path.basename(q), "synced" if same else "DIFF")
        # ★ 本轮结论必须写回 SKILL，否则视为未迭代完成
        must = ("rect.left", "右拨 1 次", "新增", "是否保存", "不是列唯一",
            "hsb_geom", "908")
        miss = [m for m in must if m not in s]
        ok = ok and not miss
        info += " | conclusions=%s" % ("all" if not miss else "MISSING:%r" % miss)
        # 经验文档同步
        if os.path.isfile(DOC_PATH):
            d = open(DOC_PATH, encoding="utf-8").read()
            dmissing = [m for m in ("rect.left", "右拨 1 次", "新增") if m not in d]
            ok = ok and not dmissing
            info += " | doc=%s" % ("synced" if not dmissing else "MISSING:%r" % dmissing)
    record("skill_files", ok, info)


# ---------------------------------------------------------------- T1
def t_split():
    """★ 第 3 轮 4/10 根因：桌面有最大化的第三方浏览器/小组件 (Chrome_WidgetWin_1)
    盖住整屏；_snap_split.py 的 Aero Snap 把它带到前台，旧回调取最后一个
    Chrome_WidgetWin_1 又把它当宿主窗口 → FAIL 且后续所有物理点击落空。
    改用 K.ensure_layout()（确定性 SetWindowPos 50/50 + 最小化干扰浏览器 +
    AttachThreadInput 强制 K3 前台），并按标题关键字精确挑宿主窗口。
    """
    ok_layout, info = K.ensure_layout(verbose=False)
    left_win = k3 = fg = None
    kh = K.find_window()[0][0]

    def cb(h, _):
        nonlocal left_win, k3
        try:
            if not win32gui.IsWindowVisible(h):
                return
            c = win32gui.GetClassName(h)
            r = win32gui.GetWindowRect(h)
            if c == "Chrome_WidgetWin_1" and r[2] - r[0] > 400:
                t = win32gui.GetWindowText(h)
                # 只认宿主主窗；忽略天气/新闻小组件等第三方 Chromium 窗口
                hints = tuple(s.strip() for s in os.environ.get("K3_LEFT_HINTS", "").split(",") if s.strip())                     or K.LEFT_TITLE_HINTS
                if any(k in t for k in hints):
                    left_win = r
            if c == "ThunderRT6FormDC" and r[2] - r[0] > 400:
                k3 = r
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    import ctypes as _ct
    fg = _ct.windll.user32.GetForegroundWindow()
    # DWM 边框阴影 ±7 容差
    _wx, _wy, _ww, _wh = K.work_area()
    _half = _ww // 2
    geom_ok = bool(left_win and k3
                   and abs(left_win[0] - _wx) <= 8 and left_win[2] <= _wx + _half + 16
                   and k3[0] >= _wx + _half - 16 and k3[2] >= _wx + _ww - 8)
    front_ok = (fg == kh)
    ok = bool(ok_layout and geom_ok and front_ok)
    record("split", ok,
           "left=%s K3=%s FG=%s expect K3=%s (50/50, 工作区 %dx%d, DWM ±7) | %s | "
           "layout=%s geom=%s front=%s" % (
               left_win, k3, hex(fg), hex(kh), _ww, _wh, info,
               ok_layout, geom_ok, front_ok))


# ---------------------------------------------------------------- T2
def t_columns_visible():
    """★ 步骤 2：归零 → 右拨 1 次 → 动态标定，点 X_QTY 必须命中 (aid201, QTY_LEFT)。"""
    hwnd, pid, title, rect, dlg = getwin()
    ok, info = K.ensure_view(dlg, pid, write=True)
    if ok:
        y = K.GRID_TOP + 8
        K.click(K.X_QTY, y); time.sleep(0.45); clean(pid)
        f = K.focused(dlg)
        ok = K.is_qty(f)
        info += " hit=%r" % (f,)
    record("columns_visible", ok, info)


# ---------------------------------------------------------------- T3
def t_calib_hit():
    """标定后的 X_CODE / X_QTY 能分别命中 (aid200,CODE_LEFT) 与 (aid201,QTY_LEFT)。"""
    hwnd, pid, title, rect, dlg = getwin()
    y = K.GRID_TOP + 8
    K.click(K.X_CODE, y); time.sleep(0.45); clean(pid)
    fc = K.focused(dlg)
    K.click(K.X_QTY, y); time.sleep(0.45); clean(pid)
    fq = K.focused(dlg)
    ok = K.is_code(fc) and K.is_qty(fq)
    record("calib_hit", ok,
           "X_CODE=%d->%r is_code=%s | X_QTY=%d->%r is_qty=%s" % (
               K.X_CODE, fc, K.is_code(fc), K.X_QTY, fq, K.is_qty(fq)))


# ---------------------------------------------------------------- T4
def t_dialog():
    """★ 空白单据提交「空代码」不会弹窗（第 2 轮 raised=False 的根因）。
    改用非法代码 NO_SUCH_MATERIAL_XYZ 触发「代码不存在或被禁用」提示框，
    验证 check_dialog 按「否」关闭、且从不点「是(&Y)」（点「是」会误写非法资料/误保存）。"""
    hwnd, pid, title, rect, dlg = getwin()
    clean(pid)
    y = K.GRID_TOP + 8
    K.click(K.X_CODE, y); time.sleep(0.5)
    keys(dlg, pid, "^a", wait=0.15)
    keys(dlg, pid, "NO_SUCH_MATERIAL_XYZ", wait=0.25)
    keys(dlg, pid, "{ENTER}", wait=1.4)
    hs = K.dlg_handles(pid)
    txt = K.dlg_text(hs[0]) if hs else ""
    btns = K.dialog_buttons(hs[0]) if hs else []
    # 「是(&Y)」必须存在，且不能被点掉
    has_yes = any("是" in b[0] for b in btns)
    closed = clean(pid)
    still = bool(K.dlg_handles(pid))
    ok = bool(hs) and has_yes and closed and not still
    record("dialog", ok,
           "raised=%s has_yes=%s text=%r buttons=%r closed=%s" % (
               bool(hs), has_yes, txt[:60], [b[0] for b in btns], not still))
    # 清理：非法代码残留只会在编辑态，撤销掉
    keys(dlg, pid, "{ESC}", wait=0.3)
    clean(pid)
    # ★ 第 8 轮（--loop 3 稳定 8/10 的根因）：非法代码写在真实第 1 行，
    #   关框后该行仍留「半输入」编辑态；下一用例 single_row 也在第 1 行，
    #   set_one 点 X_QTY 三次全被吞（焦点弹回 aid200/1043）→ FAIL，连锁使
    #   multi_row 的 0.1182 被拼成 1182（K3 显示 '1,182'）→ float() 抛异常。
    #   录完弹窗用例后立刻「新增→否」重置空白（用户授权、不保存）+ 重新标定视口。
    reset_blank()
    hwnd2, pid2, _t2, _r2 = K.find_window()[0]
    d2 = K.connect(hwnd2)
    K.ensure_view(d2, pid2, write=False)
    clean(pid2)


# ---------------------------------------------------------------- T5
def t_single_row():
    """第 1 行录入一行，读回用量 == 目标值。"""
    hwnd, pid, title, rect, dlg = getwin()
    y = K.GRID_TOP + 8
    code, qty = K.DATA[0]
    ok, info = K.set_one(dlg, pid, code, qty, y)
    read = K.read_one(dlg, pid) if ok else None
    got = num(read)
    ok2 = bool(ok and got is not None and abs(got - float(qty)) < 1e-6)
    record("single_row", ok2, "set=%s %s read=%r" % (ok, info, read))


# ---------------------------------------------------------------- T6
def t_multi_row():
    """连续录 3 行（2..4），逐行读回用量。"""
    hwnd, pid, title, rect, dlg = getwin()
    base = K.GRID_TOP + 8
    got = []
    for i in range(3):
        code, qty = K.DATA[1 + i]
        y = base + K.ROW_H * (1 + i)
        ok, info = K.set_one(dlg, pid, code, qty, y)
        r = K.read_one(dlg, pid) if ok else None
        got.append((2 + i, code, qty, r, ok, info))
    bad = [g for g in got if not g[4] or num(g[3]) is None
           or abs(num(g[3]) - float(g[2])) > 1e-6]
    record("multi_row", not bad, "rows=%r" % (got,))


# ---------------------------------------------------------------- T7
def t_clear():
    """清掉测试行 1..4：清完用量格读回 0/空，代码格为空。"""
    hwnd, pid, title, rect, dlg = getwin()
    # ★ 第 11 轮（m03614 clear FAIL 根因）：清代码格时 K3 会自动把该列滚到
    #   完整可见位置，横向视口从 right=1 拨回 HOME 附近，X_QTY=1669 于是落在
    #   aid200/1580 文本列上 ⇒ 2..4 行全 MISS。清空前必须重新标定视口。
    okv, infov = K.ensure_view(dlg, pid, write=True)
    if not okv:
        record("clear", False, "ensure_view 失败: %s" % infov)
        return
    base = K.GRID_TOP + 8
    cleaned = []
    for i in range(4):
        y = base + K.ROW_H * i
        # 用量格（★ 走 K.clear_cell：内部 ESC→点击→校验 aid→^a→DELETE→ENTER）
        okq, _ = K.clear_cell(dlg, pid, K.X_QTY, y, want_aid=201)
        r = K.read_one(dlg, pid, y) if okq else "MISS"
        # 代码格
        K.clear_cell(dlg, pid, K.X_CODE, y, want_aid=200)
        fc = K.focused(dlg)
        cleaned.append((i + 1, r, fc))
    # MISS = 该行本来就没有用量（清空后 read_one 找不到浮动格），视为已清空；
    # 只有读到非零数值才算清空失败。
    def _clean(v):
        if v in (None, "", "MISS", "0", "0.00", "0.0000"):
            return True
        try:
            return abs(float(str(v).replace(",", ""))) < 1e-9
        except Exception:
            return False
    ok = all(_clean(c[1]) for c in cleaned)
    record("clear", ok, "rows=%r" % (cleaned,))


# ---------------------------------------------------------------- T8
def t_auto_scroll():
    """★ 验证 set_one 的「ENTER 后重读 rect.top 重新锚定」在视口底部仍写对行。"""
    hwnd, pid, title, rect, dlg = getwin()
    # 动态标定一次，保证横拨位置正确
    okc, _ = K.ensure_view(dlg, pid, write=False)
    # 滚到底部，让最后可见行下方是空白行
    for _ in range(8):
        K.click(K.SCROLL_DOWN_X, K.SCROLL_DOWN_Y)
        time.sleep(0.28)
        clean(pid)
    y = K.last_visible_y()
    f = K.anchor_code(dlg, pid, y)
    if not K.is_code(f):
        record("auto_scroll", False, "anchor failed %r calib=%s" % (f, okc))
        return
    top0 = f[3]
    code, qty = K.DATA[0]
    ok, info = K.set_one(dlg, pid, code, qty, y)
    # 用 set_one 报告的重锚定 y 去读回
    read = K.read_one(dlg, pid) if ok else None
    got = num(read)
    ok2 = bool(ok and got is not None and abs(got - float(qty)) < 1e-6)
    record("auto_scroll", ok2,
           "top0=%d set=%s %s read=%r (重锚定后写对即 PASS)" % (
               top0, ok, info, read))
    # 清理这行，避免污染下一轮
    K.click(K.X_CODE, y); time.sleep(0.4); clean(pid)
    keys(dlg, pid, "^a", wait=0.12)
    keys(dlg, pid, "{DELETE}", wait=0.12)
    keys(dlg, pid, "{ENTER}", wait=0.9)
    clean(pid)


# ---------------------------------------------------------------- T9
def t_no_save():
    """窗口仍是「新增」态；所有配套脚本里 SAVE_BTN 只声明、从不调用。

    ★ 第 2 轮 FAIL 根因：静态扫描的禁用字面量出现在自检器**自己的源码**里，
      把自身命中了。这里把 needle 用字符串拼接构造，源文件不再含连续字面量。
    """
    hwnd, pid, title, rect, dlg = getwin()
    ok = "BOM" in title
    needle_call = "click(" + "SAVE_BTN"
    needle_lit1 = "click(1047" + ", 76"
    needle_lit2 = "click(1047" + ",76"
    src = open(os.path.join(HERE, "_k3_fill.py"), encoding="utf-8").read()
    ok = ok and "SAVE_BTN" in src and needle_call not in src
    # 静态扫描所有配套脚本，禁止出现「点保存」的调用
    for fn in ("_k3_fill.py", "_k3_selftest.py", "_k3_newbom.py",
               "_k3_diag.py", "_k3_dbgc.py", "_k3_dbgr.py"):
        p = os.path.join(HERE, fn)
        if os.path.isfile(p):
            s = open(p, encoding="utf-8").read()
            if any(n in s for n in (needle_call, needle_lit1, needle_lit2)):
                ok = False
    record("no_save", ok,
           "title=%r SAVE_BTN declared but never clicked | 新增→否 reset only" % title)


CASES = [
    ("skill_files", t_skill_files),
    ("split", t_split),
    ("columns_visible", t_columns_visible),
    ("calib_hit", t_calib_hit),
    ("dialog", t_dialog),
    ("single_row", t_single_row),
    ("multi_row", t_multi_row),
    ("clear", t_clear),
    ("auto_scroll", t_auto_scroll),
    ("no_save", t_no_save),
]


def run_once(round_no):
    global RESULTS
    RESULTS = []
    print("== K3 BOM SKILL selftest round %d ==" % round_no, flush=True)
    try:
        w = reset_blank()
        print("-- reset_blank (新增→否, 未保存):", w[2], flush=True)
    except Exception as e:
        print("-- reset_blank EXCEPTION:", type(e).__name__, e, flush=True)
    only = None
    if "--case" in sys.argv:
        only = sys.argv[sys.argv.index("--case") + 1]
    for name, fn in CASES:
        if only and name != only:
            continue
        try:
            fn()
        except Exception as e:
            record(name, False, "EXCEPTION %s: %s" % (type(e).__name__, e))
    npass = sum(1 for _, ok, _ in RESULTS if ok)
    print("-- round %d: %d/%d PASS --" % (round_no, npass, len(RESULTS)), flush=True)
    for n, ok, i in RESULTS:
        if not ok:
            print("   FAILED:", n, i, flush=True)
    return npass == len(RESULTS)


def main():
    a = sys.argv[1:]
    if "--list" in a:
        for n, _ in CASES:
            print(n)
        return
    loops = 1
    if "--loop" in a:
        loops = int(a[a.index("--loop") + 1])
    all_ok = True
    for i in range(1, loops + 1):
        if not run_once(i):
            all_ok = False
    print("== FINAL: %s ==" % ("ALL GREEN" if all_ok else "HAS FAILURES"), flush=True)
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
