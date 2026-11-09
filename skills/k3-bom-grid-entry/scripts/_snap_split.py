# -*- coding: utf-8 -*-
"""把「Agent 界面」贴到左半屏、金蝶 K3 贴到右半屏。

优先用 Windows 贴靠（Aero Snap：Win+Left / Win+Right）；
贴靠结果不合格时回退 SetWindowPos 精确 50/50。

用法：
    python _snap_split.py                 # 自动识别窗口
    python _snap_split.py left            # 只贴左半屏（当前前台窗口）
    python _snap_split.py right           # 只贴右半屏（当前前台窗口）
    python _snap_split.py list            # 只列出候选大窗口，不动手

工作区尺寸从系统读取（SPI_GETWORKAREA），不写死分辨率。
"""
import ctypes
import os
import sys
import time
from ctypes import wintypes as wt

import win32con
import win32gui

u32 = ctypes.windll.user32
u32.IsZoomed.argtypes = [wt.HWND]
u32.IsZoomed.restype = wt.BOOL
u32.SystemParametersInfoW.argtypes = [wt.UINT, wt.UINT, wt.POINTER(wt.RECT), wt.UINT]

VK_LWIN, VK_LEFT, VK_RIGHT, KEYUP = 0x5B, 0x25, 0x27, 0x0002
SPI_GETWORKAREA = 0x0030

# 候选窗口：Agent 宿主（Chromium 系）与 K3（VB6）
LEFT_CLASSES = ("Chrome_WidgetWin_1", "Chrome_WidgetWin_0", "MozillaWindowClass")
# 左侧宿主窗口标题关键字；可用 K3_LEFT_HINTS=foo,bar 覆盖
LEFT_TITLE_HINTS = ("宿主", "宿主")
K3_CLASSES = ("ThunderRT6FormDC", "ThunderRT6MDIForm")


def work_area():
    r = wt.RECT()
    u32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(r), 0)
    return r.left, r.top, r.right - r.left, r.bottom - r.top


SW, SH = work_area()


def key(vk, up=False):
    u32.keybd_event(vk, 0, KEYUP if up else 0, 0)
    time.sleep(0.03)


def gtitle(h):
    n = u32.GetWindowTextLengthW(h) + 2
    buf = ctypes.create_unicode_buffer(n)
    u32.GetWindowTextW(h, buf, n)
    return buf.value


def snap(hwnd, vk, tries=3):
    """用 Aero Snap 贴靠；返回贴靠后的矩形。"""
    for i in range(tries):
        u32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        before = win32gui.GetWindowRect(hwnd)
        key(VK_LWIN)
        time.sleep(0.15)
        key(vk)
        time.sleep(0.08)
        key(vk, up=True)
        time.sleep(0.15)
        key(VK_LWIN, up=True)
        time.sleep(1.0)
        after = win32gui.GetWindowRect(hwnd)
        print("  snap try%d: %s -> %s" % (i + 1, before, after), flush=True)
        if after != before:
            return after
    return win32gui.GetWindowRect(hwnd)


def big_windows(min_w=400, min_h=300):
    out = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            r = win32gui.GetWindowRect(h)
            if r[2] - r[0] > min_w and r[3] - r[1] > min_h:
                out.append((h, gtitle(h), win32gui.GetClassName(h), r))

    win32gui.EnumWindows(cb, None)
    return out


def tile(hwnd, x, y, w, h):
    if u32.IsZoomed(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.6)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h,
                          win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
    time.sleep(0.4)


def dump(wins):
    for h, t, c, r in wins:
        print("%s | %-14s | %-40s | %s" % (hex(h), c, t[:38], r), flush=True)


def half(hwnd, side):
    """side: 'left' | 'right'"""
    if side == "left":
        snap(hwnd, VK_LEFT)
        r = win32gui.GetWindowRect(hwnd)
        if r[0] <= 8 and r[2] <= SW // 2 + 16:
            return r
        tile(hwnd, 0, 0, SW // 2, SH)
    else:
        snap(hwnd, VK_RIGHT)
        r = win32gui.GetWindowRect(hwnd)
        if r[2] >= SW - 8 and r[0] >= SW // 2 - 16:
            return r
        tile(hwnd, SW // 2, 0, SW - SW // 2, SH)
    return win32gui.GetWindowRect(hwnd)


def main():
    arg = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
    wins = big_windows()

    if arg == "list":
        dump(wins)
        return

    if arg in ("left", "right"):
        h = win32gui.GetForegroundWindow()
        r = half(h, arg)
        print("%s -> %s" % (arg, r), flush=True)
        return

    print("work area: %dx%d" % (SW, SH), flush=True)
    dump(wins)

    hints = tuple(s.strip() for s in os.environ.get("K3_LEFT_HINTS", "").split(",") if s.strip())         or LEFT_TITLE_HINTS
    left = [w for w in wins if w[2] in LEFT_CLASSES or any(k in w[1] for k in hints)]
    k3 = [w for w in wins if w[2] in K3_CLASSES or w[1].startswith("BOM")]

    if not k3:
        raise SystemExit("找不到 K3 窗口（类名 ThunderRT6FormDC），请先打开「BOM单 - 新增」界面")

    kh = max(k3, key=lambda w: w[3][2] - w[3][0])[0]
    print("K3 hwnd", hex(kh), flush=True)
    print("snap K3 =", half(kh, "right"), flush=True)

    if left:
        lh = max(left, key=lambda w: w[3][2] - w[3][0])[0]
        print("left hwnd", hex(lh), flush=True)
        print("snap left =", half(lh, "left"), flush=True)
    else:
        print("未识别到 Agent 窗口，只贴了 K3。可用 `python _snap_split.py left` 手动贴左侧。")

    print("FINAL work=%dx%d" % (SW, SH), flush=True)


if __name__ == "__main__":
    main()
