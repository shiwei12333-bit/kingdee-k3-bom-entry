# 安全与隐私

本文件说明 `k3-bom-skill` 在数据层面做了什么、不做什么，以及你在公开或私有使用它时应注意什么。

## 一、运行时的网络行为

**零网络**。仓库内所有 Python 脚本（`_k3_fill.py`、`_k3_put2.py`、`_k3_put3.py`、`_k3_qty.py`、`_k3_newbom.py`、`_k3_selftest.py`、`_snap_split.py`）只使用：

- `ctypes` / `win32gui` / `win32con` / `win32process` —— Windows 本地 API
- `pywinauto` —— 本地 UIA
- `time` / `json` / `os` / `csv` / `re`

**没有任何** `requests`、`urllib`、`socket`、`http` 调用。不上报、不下载、不埋点、无遥测、无自动更新、无许可证校验回连。

可以用下面的命令自查：

```bash
grep -rn "import requests\|import urllib\|import socket\|urlopen\|http" skills/k3-bom-grid-entry/scripts/*.py
# 期望输出为空
```

## 二、磁盘写入

脚本运行时会写两个**本地**文件，都在 `scripts/` 目录下：

| 文件 | 内容 | 是否入库 |
|---|---|---|
| `_k3_calib.json` | 本机标定出来的像素坐标（列位置、滚动条位置、行高） | ❌ `.gitignore` 排除 |
| `__pycache__/` | Python 字节码缓存 | ❌ `.gitignore` 排除 |

**不写日志文件**。所有输出都只打印到终端，避免物料清单被意外落盘留存。

真实业务数据 `data.json` / `data.csv` 同样被 `.gitignore` 排除 —— 但请注意：**忽略规则只对未 `git add` 的文件生效**。提交前请始终执行：

```bash
git status --short
git diff --cached --stat
```

## 三、仓库内不含什么

本项目在公开发布前做过一次完整脱敏，仓库内**没有**：

- 真实物料编码、真实 BOM 单号
- 真实公司名、部门名、人名、邮箱
- 真实本机绝对路径（`C:\Users\<name>\...`、私有盘符路径等）
- 真实用量合计值、真实生产数据
- 任何访问令牌、Cookie、账号密码、数据库连接串

`data.example.json` / `data.example.csv` 使用的是 `1.99.001` ~ `1.99.003` 这类**合成占位编码**，`examples/bom-36rows.md` 使用 `1.99.001` ~ `1.99.036` 与合成用量（前 24 行合计 `50.25`，36 行合计 `77.25`），与任何真实 BOM 无关。

## 四、自动检查

仓库自带三道 CI 检查（`.github/workflows/lint.yml`，push / PR 时自动跑）：

| 脚本 | 作用 |
|---|---|
| `.github/scripts/check_syntax.py` | 全部 Python 脚本 `ast.parse` 语法检查 |
| `.github/scripts/check_secrets.py` | 扫描密码 / token / 私钥 / 绝对路径 / 邮箱等敏感模式 |
| `.github/scripts/check_links.py` | 检查 Markdown 内部链接与锚点是否有效 |

本地跑一遍：

```bash
python .github/scripts/check_syntax.py
python .github/scripts/check_secrets.py
python .github/scripts/check_links.py
```

`check_secrets.py` **扫描全部文本文件**（包括文档），只跳过 `.git/` 与自身。任何新增文档都在覆盖范围内。

## 五、贡献者注意

提交 PR 前请确认：

1. 没有把 `data.json` / `data.csv` / `_k3_calib.json` 加进暂存区；
2. 截图里没有露出真实物料编码、客户名、公司名（**脱敏或裁掉**，或改用文字描述）；
3. 贴的报错信息里没有真实路径与账号；
4. 新增的示例数据一律使用 `1.99.xxx` 风格的合成编码。

## 六、报告问题

发现仓库里仍残留敏感信息（真实路径、真实编码、凭据），请开一个 Issue 说明位置（**不要在 Issue 里复述敏感内容本身**），或直接提 PR 修正。

## 七、责任边界

本项目的完整免责声明见 [README 的「免责与边界」章节](README.md#免责与边界)。一句话概括：**脚本负责敲键盘，判断力仍然是你的。**
