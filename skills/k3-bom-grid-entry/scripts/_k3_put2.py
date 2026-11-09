# -*- coding: utf-8 -*-
"""连续录入（第 13 轮修正版）：绝不 ESC、绝不读回、每格 click→^a→DELETE→输入→ENTER。

用法:
  python _k3_put2.py code <from> <to>      # 商品代码列，行号从 1 起
  python _k3_put2.py qty  <from> <to>      # 用量列
  python _k3_put2.py coderaw 1.99.001 2  # 单条直接写某行
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import _k3_fill as K

if not K.DATA:
    print("ERROR: 没有可录入的数据。请提供 data.json / data.csv，或设置环境变量 K3_DATA 指向数据文件。")
    print("       格式见 data.example.json（[[\"物料代码\", \"用量\"], ...]）")
    raise SystemExit(2)

X_CODE = 1090
X_QTY = 1400   # ★ 用量列；1742~1800 是「备注」
Y1 = 496
ROW_H = 16
PAGE_ROWS = 24

CODES = [c for c, q in K.DATA]
QTYS = [q for c, q in K.DATA]


def esc(s):
    return str(s).replace(".", "{.}")


def put(dlg, pid, x, y, text):
    """一次写入。返回 True 表示已发送完（不做读回校验）。"""
    K.click(x, y)
    time.sleep(0.18)
    K.safe_keys(dlg, pid, "^a", wait=0.12)
    K.safe_keys(dlg, pid, "{DELETE}", wait=0.12)
    K.safe_keys(dlg, pid, esc(text), wait=0.12)
    K.safe_keys(dlg, pid, "{ENTER}", wait=0.75)
    d = K.check_dialog(pid, close=True)
    if d:
        print("   ! dialog: %s" % d.replace("\r\n", " ")[:70], flush=True)
        return False
    return True


def run(mode, lo, hi):
    hwnd, pid, title, rect = K.find_window()[0]
    dlg = K.connect(hwnd)
    print("win:", title, pid, flush=True)
    x = X_CODE if mode == "code" else X_QTY
    src = CODES if mode == "code" else QTYS
    for row in range(lo, hi + 1):
        y = Y1 + (row - 1) * ROW_H
        val = src[row - 1]
        ok = put(dlg, pid, x, y, val)
        print("row %2d y=%3d %-16s %s" % (row, y, val, "OK" if ok else "FAIL"), flush=True)
        if not ok:
            print("ABORT at row", row, flush=True)
            return 1
    print("DONE %s rows %d..%d" % (mode, lo, hi), flush=True)
    return 0


def main():
    m = sys.argv[1]
    if m == "coderaw":
        code = sys.argv[2]
        row = int(sys.argv[3])
        hwnd, pid, title, rect = K.find_window()[0]
        dlg = K.connect(hwnd)
        y = Y1 + (row - 1) * ROW_H
        put(dlg, pid, X_CODE, y, code)
        print("raw put row", row)
        return
    lo = int(sys.argv[2])
    hi = int(sys.argv[3])
    run(m, lo, hi)


main()
