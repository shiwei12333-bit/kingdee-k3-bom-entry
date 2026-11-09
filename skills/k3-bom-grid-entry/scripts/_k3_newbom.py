# -*- coding: utf-8 -*-
"""探测/执行「新增」重置：点工具栏「新」，读弹窗按钮，按「否」放弃保存。

用法：
  python _k3_newbom.py probe    # 只点新增并打印弹窗结构，不关
  python _k3_newbom.py reset    # 点新增，遇弹窗点「否」，确认明细归零
严禁点击保存 (1047,76)。
"""
import sys, time
import win32gui
import _k3_fill as K

NEW_BTN = (985, 78)     # 工具栏「新」（OCR: 新 x=979,y=75）
NO_HINT = ("否", "No", "N", "取消")


def dump_dialogs(pid):
    out = []
    for h in K.dlg_handles(pid):
        r = win32gui.GetWindowRect(h)
        btns = []

        def cb(c, _):
            try:
                t = win32gui.GetWindowText(c)
                cr = win32gui.GetWindowRect(c)
                cls = win32gui.GetClassName(c)
                if t:
                    btns.append((t, cls, cr))
            except Exception:
                pass
        win32gui.EnumChildWindows(h, cb, None)
        out.append((h, r, btns))
    return out


def click_no(pid, verbose=True):
    """点弹窗里的「否/取消」按钮（放弃保存）；返回 (clicked, 描述)。"""
    for h, r, btns in dump_dialogs(pid):
        for t, cls, cr in btns:
            if any(k in t for k in NO_HINT):
                cx = (cr[0] + cr[2]) // 2
                cy = (cr[1] + cr[3]) // 2
                K.click(cx, cy)
                time.sleep(0.6)
                if verbose:
                    print("clicked NO:", repr(t), cls, cr, flush=True)
                return True, t
    # 兜底：右下角（K3 的「否」通常在右下）
    if dump_dialogs(pid):
        h, r, btns = dump_dialogs(pid)[0]
        K.click(r[2] - 90, r[3] - 35)
        time.sleep(0.6)
        if verbose:
            print("clicked fallback bottom-right", r, flush=True)
        return True, "fallback"
    return False, "no dialog"


def reset(verbose=True):
    hwnd, pid, title, rect = K.find_window()[0]
    K.clean_dialogs(pid)
    if verbose:
        print("before:", title, flush=True)
    K.click(*NEW_BTN)
    time.sleep(1.0)
    ds = dump_dialogs(pid)
    if verbose:
        for h, r, btns in ds:
            print("dialog", hex(h), r, flush=True)
            for t, cls, cr in btns:
                print("   ", repr(t), cls, cr, flush=True)
    if ds:
        click_no(pid, verbose)
        time.sleep(0.8)
    # 可能还有二次确认
    for _ in range(3):
        if not K.dlg_handles(pid):
            break
        click_no(pid, verbose)
        time.sleep(0.6)
    K.clean_dialogs(pid)
    after = K.find_window()[0]
    if verbose:
        print("after:", after[2], flush=True)
    return after


if __name__ == "__main__":
    a = sys.argv[1:]
    if a and a[0] == "probe":
        reset()
    else:
        reset()
        hwnd, pid, title, rect = K.find_window()[0]
        dlg = K.connect(hwnd)
        print("focused:", K.focused(dlg), flush=True)
        print("calib:", K.X_CODE, K.X_QTY, K.CODE_LEFT, K.QTY_LEFT, flush=True)
