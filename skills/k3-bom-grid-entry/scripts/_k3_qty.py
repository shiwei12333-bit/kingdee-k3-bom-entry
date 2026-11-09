# -*- coding: utf-8 -*-
"""用量列录入（第 14 轮）：点格后先校验焦点确实是 aid201 用量格，再输入。

用法:
  python _k3_qty.py <from> <to> <base_row> <base_y>
  python _k3_qty.py probe          # 只探测当前视口用量列坐标
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import _k3_fill as K

if not K.DATA:
    print("ERROR: 没有可录入的数据。请提供 data.json / data.csv，或设置环境变量 K3_DATA 指向数据文件。")
    print("       格式见 data.example.json（[[\"物料代码\", \"用量\"], ...]）")
    raise SystemExit(2)

X_QTY = 1400
ROW_H = 16
QTYS = [q for c, q in K.DATA]
CAND = [1400, 1380, 1420, 1360, 1440, 1340, 1460]


def esc(s):
    return str(s).replace(".", "{.}")


def probe_y(dlg, pid, y):
    """在 y 行横扫找 aid201 且 left 最小的用量格，返回 (x, left) 或 None。"""
    best = None
    for x in CAND:
        K.click(x, y)
        time.sleep(0.18)
        f = K.focused(dlg)
        if f and f[0] == "201":
            left = f[4]
            if best is None or left < best[1]:
                best = (x, left)
        K.check_dialog(pid, close=True)
    return best


def put_qty(dlg, pid, y, val, x):
    K.click(x, y)
    time.sleep(0.2)
    f = K.focused(dlg)
    if not f or f[0] != "201":
        return False, "焦点不是用量格 f=%r" % (f,)
    K.safe_keys(dlg, pid, "^a", wait=0.12)
    K.safe_keys(dlg, pid, "{DELETE}", wait=0.12)
    K.safe_keys(dlg, pid, esc(val), wait=0.15)
    f2 = K.focused(dlg)
    got = f2[1] if f2 else ""
    K.safe_keys(dlg, pid, "{ENTER}", wait=0.7)
    d = K.check_dialog(pid, close=True)
    if d:
        return False, "弹窗:%s" % d.replace("\r\n", " ")[:60]
    try:
        ok = abs(float(got.replace(",", "")) - float(val)) < 1e-9
    except ValueError:
        ok = False
    return ok, got


def main():
    hwnd, pid, title, rect = K.find_window()[0]
    dlg = K.connect(hwnd)
    print("win:", title, pid, flush=True)
    if sys.argv[1] == "probe":
        for y in (496, 512, 528):
            print("y=%d -> %r" % (y, probe_y(dlg, pid, y)), flush=True)
        return
    lo, hi = int(sys.argv[2]), int(sys.argv[3])
    base_row, base_y = int(sys.argv[4]), int(sys.argv[5])
    x = X_QTY
    for row in range(lo, hi + 1):
        y = base_y + (row - base_row) * ROW_H
        val = QTYS[row - 1]
        ok, info = put_qty(dlg, pid, y, val, x)
        print("row %2d y=%3d qty=%-10s %s %s" % (row, y, val, "OK" if ok else "FAIL", info), flush=True)
        if not ok:
            print("ABORT at row", row, flush=True)
            return
    print("DONE qty rows %d..%d" % (lo, hi), flush=True)


main()
