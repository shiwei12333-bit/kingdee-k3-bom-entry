---
name: k3-bom-grid-entry
version: 1.0.0
description: 在金蝶 K3「BOM单 - 新增/修改」的 SPR32A70 明细表格里批量录入商品代码与用量：先把操作端窗口与 K3 平均左右分屏、横拨一次让「用量」列进入视口，然后按页（每页可见行数-1）用固定坐标 click→全选→删除→输入→ENTER 连续录入，先填完全部代码再回头按页填全部用量，每页填完截图确认后才滚动，最后按合计值校验，全程不点保存。
license: MIT
tags:
  - kingdee
  - k3
  - erp
  - bom
  - ui-automation
  - pywinauto
  - windows
  - rpa
triggers:
  - 金蝶 K3 录入 BOM 明细
  - ERP 表格自动填
  - K3 弹窗卡住
  - 格子输入错行
  - K3 看不到用量列
  - 批量录物料号
requires:
  os: windows
  python: ">=3.9"
  packages:
    - pywinauto>=0.6.8
    - pywin32>=306
    - comtypes>=1.2.0
metadata:
  author: YG
  verified_on:
    kingdee_version: K3 WISE / K3 Express (KDSYSTEM)
    grid_control: SPR32A70
    screen: 1920x1080, taskbar 48px
---

# 金蝶 K3 BOM 明细表批量录入

把一列「物料号 + 用量」静默录入已打开的 K3 `BOM单 - 新增` 界面：**只填商品代码与用量两列**，其余列（商品名称、规格型号、商品属性、单位、子项类型、使用状态、发料仓库/仓位）由 K3 在提交代码后自动带出，**结束时绝不点保存**。

同一套方法适用于 K3 任意 `SPR32A70` 表格控件：销售订单明细、生产任务单、调拨单等。

- 本流程在 36 行 BOM 上实测一次通过，合计 `77.25` 与源 Excel 完全一致；
- 前 12 轮失败的三个根因见第 1 节，务必先读；
- 全程通过 `mcp`/`pywinauto` 的 UIA 与屏幕坐标混合操作，无需 K3 提供任何接口。

---

## 目录

1. [三个致命根因](#1-三个致命根因)
2. [铁律](#2-铁律)
3. [环境与前置](#3-环境与前置)
4. [步骤 1：分屏（K3 占右半侧）](#4-步骤-1分屏k3-占右半侧)
5. [步骤 2：横拨，让「用量」列进视口](#5-步骤-2横拨让用量列进视口)
6. [步骤 3：标定列坐标、首行 y、行距、每页行数](#6-步骤-3标定列坐标首行-y行距每页行数)
7. [步骤 4：单格写入的标准动作](#7-步骤-4单格写入的标准动作)
8. [步骤 5：分页填完全部代码](#8-步骤-5分页填完全部代码)
9. [步骤 6：回到第 1 行填全部用量](#9-步骤-6回到第-1-行填全部用量)
10. [步骤 7：全表核对](#10-步骤-7全表核对)
11. [步骤 8：收尾](#11-步骤-8收尾)
12. [脚本清单](#12-脚本清单)
13. [排查顺序](#13-排查顺序)
14. [已知限制](#14-已知限制)

---

## 1. 三个致命根因

前 12 轮跑不通，全部错误都归到这三条：

| 错误做法 | 真实后果 | 正确做法 |
|---|---|---|
| 每格先发 `ESC` 退编辑态 | **ESC 直接取消该格输入**，写过的内容全清空 | 永不发 ESC；退出编辑态靠点列标题空白 `(1890, 476)` 或点下一格 |
| 用 `focused()` 读回做校验 | 无弹窗时读到的是 **UIA 陈旧缓存**，假报失败，也假报命中错误列 | 不读回；正确性靠**每页截图 + 全表合计值** |
| 滚动条坐标写死 `y=917`、`(1910,907)` | 落点在滚动条**边框之外**，箭头从未被点中 | 横向中心 y=**908**、下滚 **(1908,891)**、上滚 **(1908,468)**，并每次从 UIA `ScrollBar` 矩形重算 |

另有两条必须遵守的细节：

- `^a` 之后**必须补 `{DELETE}`**，否则 K3 把新值拼到旧值前面（`写1→11.0000`、`写0.1182→11,821.0000`）；
- 小数点必须转义：`str(qty).replace(".", "{.}")`，裸 `.` 会被 pywinauto 当特殊键吞掉（`0.1182→01182`）。

---

## 2. 铁律

0. **启动先做两件事**：① 平均左右分屏（操作端窗口在左、K3 在右）；② 横滑条右拨 1 次，直到列标题出现「用 量」。
1. **不用方向键导航**：`{DOWN}`/`{RIGHT}` 会让 K3 弹框并跳行。滚动只允许点横/竖滚动条箭头。
2. **绝不点保存**（工具栏「保存」`x≈1041~1053, y≈76`）。误点等于写库。
3. **不用剪贴板批量粘贴**：K3 会把整段 TSV 塞进单格，然后弹「代码在数据库中不存在或被禁用」。
4. **窗口尺寸一变化，所有坐标立刻作废**：重新按步骤 3 标定，禁止沿用上一屏常量。
5. **每填完一页必须截图确认，才允许滚动**。滚动后行号与 y 的对应关系全变，必须重新确认首行。
5b. **写用量前必须先把焦点固定到「用量」第 1 行格子上**（`focused()` 返回 `aid==201` 才算成功），否则会写错列；每换一页都要重新固定。
6. **一个动作后立刻检测弹窗**。弹窗期间 `type_keys` 抛 `ElementNotEnabled`，看着像按键失灵，实际是被 `#32770` 挡住。关掉弹窗后必须重新锚定，不复用旧坐标。
7. **先填完全部代码，再回头统一填用量**。逐行 code+qty 会被 K3 提交后的自动上滚 16px 打乱行号。

---

## 3. 环境与前置

- Windows 10/11，Python ≥ 3.9；
- `pip install pywinauto pywin32 comtypes`；
- 屏幕 1920×1080、任务栏 48px（本文所有坐标基于此；其他分辨率须重新标定）；
- K3 已打开一张 `BOM单 - 新增` 单据，明细区可见。

### 待录入数据从哪来

仓库里**不含任何真实物料数据**。数据按以下顺序加载（`_k3_fill.load_data()`）：

1. 环境变量 `K3_DATA` 指向的文件；
2. `scripts/data.json`：`[["物料代码", "用量"], ...]`（见 `data.example.json`）；
3. `scripts/data.csv`：每行 `代码,用量`（见 `data.example.csv`）。

加载后自动算出 `TOTAL_EXPECT`（Σ用量），用于第 10 节的合计校验。三者都没有时脚本直接退出并提示。

两个 Python 陷阱：

```python
# ❌ win32gui.GetWindowThreadProcessId / win32gui.IsZoomed 不存在
# ✅
import win32process, ctypes
pid = win32process.GetWindowThreadProcessId(hwnd)[1]
zoomed = ctypes.windll.user32.IsZoomed(hwnd)
```

- 若通过 bash 调 Python：`Set-Clipboard`、`$var=(...)` 是 PowerShell 语法，会直接报错；中文输出需 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")`。
- 定位窗口**不要靠标题**（K3 标题经 GBK 会乱码），按类名：操作端窗口 `Chrome_WidgetWin_1`、K3 `ThunderRT6FormDC`，再筛可见 + 同进程 + 最宽。
- `find_window()` 不接受 `verbose` 参数（会 TypeError）。

UIA 现实：K3 **只把当前获得焦点的单元格暴露给 UIA**，其余单元格不可枚举；`SPR32A70_SpreadSheet` 不支持 Scroll/Grid/Value/Table 任何 Pattern。判定「点中了哪个格」只能读 `descendants()` 里 `has_keyboard_focus()` 元素的 `(automation_id, window_text, class_name, rect.top, rect.left)`。

| 列 | automation_id | class |
|---|---|---|
| 商品代码 | **200** | `SPR32A70EditHScroll` |
| 用量 | **201** | `SPR32A70Float` |

> **aid 不唯一**：多个文本列都是 200，**用量与损耗率都是 201**。必须用 `(aid, rect.left)` 双条件判列：商品代码取 `aid=200` 中 `rect.left` 最小者，用量取 `aid=201` 中 `rect.left` 最小者。

---

## 4. 步骤 1：分屏（K3 占右半侧）

```bash
python scripts/_snap_split.py
```

或用底座库校验当前布局：

```python
import _k3_fill as K
K.ensure_layout()   # -> (True, '操作端窗口=(0, 0, 960, 1032) K3=(960, 0, 1920, 1032) FG=0x18a117c')
```

- 目标：左窗口 `(0,0,960,1032)`、K3 `(960,0,1920,1032)`。
- 贴靠用 Aero Snap（`Win+←/→`）：先 `SetForegroundWindow` 再发键，中间 `sleep 0.5`，否则会贴到别的窗口上。
- Aero Snap 对已贴边的窗口发 `Win+→` 无动作；不合格就回退 `SetWindowPos` 精确 50/50。
- K3 贴靠后是「最大化」态（`IsZoomed` 为真），改大小前要 `ShowWindow(SW_RESTORE)`。

---

## 5. 步骤 2：横拨，让「用量」列进视口

分屏后 K3 只有 960px 宽，「用量」默认在视口右缘之外，只能看到「使用状态」。**必须右拨一次**。

| 控件 | 960 分屏下的坐标 | 说明 |
|---|---|---|
| 横向滚动条整体 | UIA `ScrollBar rect=(977,900,924,17)` → **中心 y=908** | 唯一权威，每次从 UIA 重读 |
| 左箭头（归零） | `(987, 908)` | 点 12 次强制回 HOME |
| 右箭头（右拨） | `(1891, 908)` | **只点 1 次**，点 4 次会拨过头把用量列移出视口 |
| 列标题空白（退编辑态） | `(1890, 476)` | 不碰任何数据 |

```python
def hsb_geom(dlg=None):
    """(left_x, right_x, y)。★ 只按 class_name()=="ScrollBar" 匹配，
    用 control_type='Pane' 会一个都取不到、永远走 fallback。"""
    fb = (987, 1891, 908)
    best = None
    for c in dlg.descendants():
        if c.class_name() != "ScrollBar":
            continue
        r = c.rectangle(); w, h = r.right - r.left, r.bottom - r.top
        if h <= 30 and w > 200 and (best is None or w > best[2] - best[0]):
            best = (r.left, r.top, r.right, r.bottom)
    return (best[0] + 10, best[2] - 10, (best[1] + best[3]) // 2) if best else fb

def vsb_geom(dlg=None):
    """(up_x, up_y, down_x, down_y)。竖条 UIA 实测 rect=(1900,458,17,443)。"""
    fb = (1908, 468, 1908, 891)
    best = None
    for c in dlg.descendants():
        if c.class_name() != "ScrollBar":
            continue
        r = c.rectangle(); w, h = r.right - r.left, r.bottom - r.top
        if w <= 30 and h > 100 and (best is None or h > best[3] - best[1]):
            best = (r.left, r.top, r.right, r.bottom)
    if best:
        cx = (best[0] + best[2]) // 2
        return (cx, best[1] + 10, cx, best[3] - 10)
    return fb
```

---

## 6. 步骤 3：标定列坐标、首行 y、行距、每页行数

**操作**：先截图，看列标题行（y≈469~472）确认「用 量」列位置；再点第一行的代码格与用量格，读 `focused()` 的 `rect.left` 反推列 x。

实测（960 分屏 + 右拨 1 次 + 商品名称列被手动拉宽）：

| 项 | 值 |
|---|---|
| 表头 y | **470** |
| 第 1 数据行 y | **496** |
| 行距 `ROW_H` | **16 px** |
| 可见行数 | 25 → **每页填 24 行（可见行数 − 1）** |
| 商品代码列 x | **1090**（控件 `aid=200 left=1043`） |
| 用量列 x | **1400**（控件 `aid=201 left=1392`） |
| 表格区域 | y 461..924，右侧竖条 x≈1908，底部横条 y=908 |

列序（右拨 1 次后）：`序号 999 / 商品代码 1074 / 商品属性 / 辅助属性 / 单位 / 子类 / 用量 1394 / 损耗率 1467 / 使用状态 1560 / 发料仓库 1619 / 发料仓位 1689 / 备注 1742 / 存在替代 1860`。

> **用量列的致命混淆**：`x=1760/1800` 落在「备注」列（`focused()` 会假报 `aid=201 left=1729` —— 陈旧缓存），`x=1830/1860` 落在「损耗率」(`left=1451`)。**只有 x≈1400 是「用量」**。
> 列宽可被手动拖动，商品名称/规格型号一拉宽，后面所有列整体右移 ⇒ **每屏都要重新确认**，不能抄旧值。

---

## 7. 步骤 4：单格写入的标准动作

```python
def put(dlg, pid, x, y, text):
    click(x, y)                                            # ① 点中格子
    time.sleep(0.18)
    safe_keys(dlg, pid, "^a")                              # ② 全选
    safe_keys(dlg, pid, "{DELETE}")                        # ③ 必须在 ^a 之后，否则拼旧值
    safe_keys(dlg, pid, str(text).replace(".", "{.}"))     # ④ 小数点转义
    safe_keys(dlg, pid, "{ENTER}")                         # ⑤ 提交，否则下一次 click 丢数据
    time.sleep(0.75)
    return check_dialog(pid, close=True)                   # ⑥ 关弹窗（None 表示无弹窗）
```

- `safe_keys` = 内部捕获 `ElementNotEnabled` → `clean_dialogs(pid)` → 重试。
- 提交代码后 K3 自动带出：商品名称、规格型号、商品属性（外购）、单位（PCS）、子项类型、使用状态、发料仓库/仓位。**不要手填这些列**。
- 点空白行提交空代码会弹「金蝶提示 / 请输入 物料代码」，唯一按钮「确定(O)」中心 `(1100,605)`；弹窗期间主窗口 `isEnabled=false`，所有点击被吞。

常用弹窗：

| 弹窗文字 | 触发场景 | 处理 |
|---|---|---|
| `请输入 物料代码` + `确定(&O)` | 点了空白行 | 点 `(1100,605)` |
| `输入的代码:[xxx]在数据库中不存在或被禁用` + `是/否` | 代码非法 / 批量粘贴 / 错行 | 点「否」 |
| `当前单据已经改变，是否保存单据？` + `是/否/取消` | 点工具栏「新增」 | **永远点「否」**（`否(N)` 中心 `(1100,605)`，但一律按按钮文字点） |
| `将撤销当前改动，是否需要恢复?` + `是/否` | `Ctrl+Z` | 点「否」**不会恢复**已丢的数据 |

---

## 8. 步骤 5：分页填完全部代码（先只填代码）

```
第 1 页：python scripts/_k3_put2.py code 1 24          # 24 行
         → 截图确认第 1..24 行代码齐全、第 25 行空白
第 2 页：下滚 20 次（点 (1908,891)）                    # 视口首行变成序号 21
         → 截图确认首行序号
         python scripts/_k3_put3.py code 25 36 21 496  # base_row=21 base_y=496，写 12 行
```

- `_k3_put2.py`：常量 `Y1=496, ROW_H=16`，`y = Y1 + (row-1)*16`，适用于**表格位于最顶端**时。
- `_k3_put3.py`：`code <from> <to> <base_row> <base_y>`，`y = base_y + (row-base_row)*16`，滚动后必须用它。
- 滚动动作：`for _ in range(20): click(1908,891); sleep(0.09)` → `sleep(0.9)` → `check_dialog`。每次下滚 20 下约前进 20 行。
- 填完一页**必须截图**确认正确才允许滚动（铁律 5）。
- 实测 36 条代码：第 1..24 行一次过；下滚后首行=序号 21；第 25..36 行沿用 `base_row=21` 一次过；第 37 行起空白。

---

## 9. 步骤 6：回到第 1 行填全部用量

**填用量前先点上滚箭头回到序号 1。**

★★ **第一步必须先把焦点固定到「用量」第 1 行格子上，再开始批量输入**，否则整片用量会写错列/错行：

```python
click(X_QTY, Y1)              # X_QTY=1400, Y1=496
f = focused(dlg)              # 必须 (aid=="201", left 最小 = 用量列)
assert f[0] == "201", "焦点没固定在用量第一行，禁止开写：%r" % (f,)
```

- 焦点没固定就直接批量发按键，K3 会把值写进当前残留焦点的格子（实测会落到商品代码列/备注列）。
- 若 `aid != 201`：先 `hscroll_home` 再 `hscroll_right` 一次，重新确认 `X_QTY`，再固定焦点。
- 滚动换页后**同样要重新固定一次焦点**，不能沿用上一页的。

```
上滚：for _ in range(30): click(1908,468); sleep(0.09) → sleep(0.9)
截图确认首行是序号 1
固定焦点：click(1400, 496) + 校验 aid==201
第 1 页：python scripts/_k3_qty.py qty 1 24 1 496       # 24 行
         → 截图确认，视口合计应为 50.25
下滚 20 次 → 首行序号 21 → 重新固定焦点
第 2 页：python scripts/_k3_qty.py qty 25 36 21 496     # 12 行
         → 截图确认，全表合计应为 77.25
```

`_k3_qty.py` 的 `put_qty()` 会先 `click` 再 `focused()` 校验 `aid=="201"` 才输入（不合格返回 `焦点不是用量格`），并在写入后**立即**用 `focused()` 读一次值比对。

- 这个「写入后立即读」只在**同一次编辑态内**可靠；一旦 ENTER 提交，`focused()` 就变陈旧缓存，**不能用于最终校验**（最终校验看截图 + 合计）。

---

## 10. 步骤 7：全表核对

1. **截图对照**：按页截图，与源 Excel 逐行核对物料号 + 用量。
2. **合计值校验（最省事、最可靠）**：表底合计行 = Σ用量。36 行示例正确合计 **`77.25`**（前 24 行 `50.25`）。
3. **差值定位**：合计差多少 → 找哪一行的用量等于该差值。

---

## 11. 步骤 8：收尾

- **绝不点保存**（`x≈1047, y≈76`），保持「状态：未使用」。
- 需要清空整单重来：点工具栏「新增」→ 弹「当前单据已经改变，是否保存单据？」→ 点「否」（`scripts/_k3_newbom.py reset`）。
- 放弃全部内容：直接关窗 → 问是否保存 → 点「否」。

---

## 12. 脚本清单

```
python scripts/_snap_split.py                      # 步骤 1：分屏（操作端窗口 左 / K3 右）
python scripts/_k3_fill.py probe                   # 窗口 / 焦点 / 弹窗状态
python scripts/_k3_fill.py calib                   # 归零 + 右拨 1 次 + 动态标定列坐标
python scripts/_k3_put2.py code 1 24               # 第 1 页代码（表格在顶端时）
python scripts/_k3_put3.py code 25 36 21 496       # 滚动后代码（base_row/base_y 锚定）
python scripts/_k3_qty.py  qty 1 24 1 496          # 第 1 页用量
python scripts/_k3_qty.py  probe                   # 只探测用量列坐标
python scripts/_k3_put2.py coderaw 1.99.001 2    # 单条直写某行
python scripts/_k3_fill.py scroll up|down <N>      # 竖滚动条
python scripts/_k3_fill.py hscroll left|right <N>  # 横滚动条
python scripts/_k3_newbom.py reset                 # 新增→否，重置空白（不保存）
python scripts/_k3_selftest.py --loop 3            # 循环自检（10 个用例）
```

| 文件 | 作用 |
|---|---|
| `_k3_fill.py` | 底座库：窗口定位、分屏、点击、`safe_keys`、弹窗处理、滚动几何、标定、`ensure_view`、`clear_cell` |
| `_k3_put2.py` | 表格在顶端时的分页录入（代码/用量通用） |
| `_k3_put3.py` | 滚动后按 `base_row/base_y` 锚定录入 |
| `_k3_qty.py` | 用量专用：点格后先校验 `aid==201` 再输入，支持 probe |
| `_k3_newbom.py` | 新增→否，重置空白明细 |
| `_k3_selftest.py` | 循环自检器，10 个用例，验证可重复性 |
| `_snap_split.py` | Aero Snap 50/50 分屏 |
| `_k3_calib.json` | 列坐标标定缓存（**本机生成，已 gitignore**；首次运行自动创建） |
| `_k3_calib.example.json` | 标定文件模板，仅供参考，勿直接抄坐标 |
| `data.example.json` / `data.example.csv` | 待录入数据模板，复制为 `data.json` / `data.csv` 后填入自己的数据 |

---

## 13. 排查顺序

0. **点滚动条没反应** → `print(hsb_geom(dlg), vsb_geom(dlg))` 必须落在 UIA `ScrollBar` 矩形内（横 `(987,1891,908)`、竖 `(1908,468,1908,891)`）。写死 `917`/`907` 全在条外。
1. **看不见「用量」列**，或点用量格命中备注/损耗率 → 右拨次数不对（只拨 1 次），或列宽被拖宽需重新标定。别只看 `aid=201`，要 `(aid, rect.left)`。
2. **读回是拼接值**（`写1→11.0000`）→ `^a` 后漏了 `{DELETE}`。
3. **小数变千分位**（`0.1182→1,182`）→ 小数点没转义成 `{.}`。
4. **`type_keys` 抛 `ElementNotEnabled`** → 有弹窗，先 `check_dialog(pid, close=True)`。
5. **数据写到错行** → ENTER 后视口自滚 16px，沿用旧 y；改用 `_k3_put3.py` 的 base 锚定，或滚动后重新确认首行。
6. **某行整段消失** → 用了 ESC 退编辑态（ESC = 取消输入）。
7. **合计对不上** → 差值定位，先清后补。

---

## 14. 已知限制

- 坐标与 DPI/分辨率/任务栏高度/列宽强相关，换机器必须重标定（步骤 3）。
- 依赖 K3 只暴露焦点单元格这一特性，无法程序化读全表，因此**校验只能靠截图 + 合计值**。
- 自动上滚 16px 的触发条件：目标行是当前可视区最后一行 + 下一行是空白 + 用 `{ENTER}` 提交。三条同时满足才出现。
- Windows 上的 UIA 缓存行为可能随 K3 小版本变化，出现「读回突然变假」时先怀疑缓存，不要怀疑列宽。

---

## 许可

MIT，见仓库根目录的 [LICENSE](../../LICENSE)。

## 安全边界（与 README 保持一致）

- **零网络**：脚本只用 Windows 本地 API 与 pywinauto，不发起任何 HTTP / socket 请求。
- **不写日志**：只打印终端，不落盘，避免物料清单外泄。
- **不点保存**：整条流程的唯一安全边界，任何脚本都不得调用保存。
- **不读数据库**：全部通过 UI 层输入，不接触 K3 的表结构。
- 完整说明见仓库根目录 `SECURITY.md`；操作前检查单见 `docs/CHECKLIST.md`；
  换机器标定见 `docs/CALIBRATION.md`。

## 与其他文档的关系

| 文件 | 作用 |
|---|---|
| `SKILL.md`（本文） | Agent 执行手册：按什么顺序做、每步怎么验证 |
| `README.md` | 人看的入口：这是什么、怎么装、怎么跑 |
| `docs/CALIBRATION.md` | 换环境后怎么重新取坐标 |
| `docs/CHECKLIST.md` | 每次实操照着打勾 |
| `docs/TROUBLESHOOTING.md` | 踩坑全记录与排查顺序 |
