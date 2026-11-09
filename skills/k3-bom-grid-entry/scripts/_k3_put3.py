# -*- coding: utf-8 -*-
"""连续录入（第 13 轮修正版，带视口基准）：绝不 ESC、绝不读回。

用法:
  python _k3_put3.py code <from> <to> <base_row> <base_y>
  python _k3_put3.py qty  <from> <to> <base_row> <base_y>
其中 base_row 是当前视口最上面那行的序号，base_y 是它的屏幕 y。
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
ROW_H = 16

CODES = [c for c, q in K.DATA]
QTYS = [q for c, q in K.DATA]


def esc(s):
    return str(s).replace(".", "{.}")


def put(dlg, pid, x, y, text):
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


def main():
    mode = sys.argv[1]
    lo = int(sys.argv[2])
    hi = int(sys.argv[3])
    base_row = int(sys.argv[4])
    base_y = int(sys.argv[5])
    hwnd, pid, title, rect = K.find_window()[0]
    dlg = K.connect(hwnd)
    print("win:", title, pid, "base_row", base_row, "base_y", base_y, flush=True)
    x = X_CODE if mode == "code" else X_QTY
    src = CODES if mode == "code" else QTYS
    for row in range(lo, hi + 1):
        y = base_y + (row - base_row) * ROW_H
        val = src[row - 1]
        ok = put(dlg, pid, x, y, val)
        print("row %2d y=%3d %-16s %s" % (row, y, val, "OK" if ok else "FAIL"), flush=True)
        if not ok:
            print("ABORT at row", row, flush=True)
            return
    print("DONE %s rows %d..%d" % (mode, lo, hi), flush=True)


main()
