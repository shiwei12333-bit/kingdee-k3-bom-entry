# 贡献指南

感谢你愿意把这个技能包做得更稳。提交前请先读完本页。

## 可以提什么

| 类型 | 说明 |
|---|---|
| 🐞 新的踩坑记录 | 换分辨率、换 K3 版本、换 DPI 缩放后遇到的问题与解法，最有价值 |
| 📐 标定参数 | 不同分辨率 / 缩放比例下 `python _k3_fill.py calib` 的实测输出 |
| 🧩 新脚本 | 能复用 `_k3_fill.py` 底座库的通用能力（如导出、核对、批量清空） |
| 📝 文档修正 | 描述不准、步骤缺失、命令跑不通 |

## 不可以提什么

- ❌ 真实物料代码、真实客户名称、真实单据号、真实价格、真实 BOM 结构
- ❌ 任何账号、密码、token、内网地址、服务器名
- ❌ 采集/上传业务数据的代码
- ❌ 绕过权限、调用 K3 未公开接口、直接读写数据库的实现

**提交前请自查**：

```bash
grep -rniE "passwd|password|token|secret|api[_-]?key|corp|内网" .
```

## 开发约定

1. Python 3.9+，只用标准库 + `pywinauto` / `pywin32` / `comtypes`，不引入新的重依赖；
2. 所有脚本保持可 `python <script>.py` 直接运行，并在文件头写清 CLI 用法；
3. **不得内置业务数据**——一律走 `data.json` / `data.csv` / `K3_DATA`；
4. **不得硬编码坐标作为唯一来源**——优先从 UIA 动态读取，硬编码只作为 fallback；
5. 提交前跑一遍语法检查：

```bash
python -c "import ast,glob;[ast.parse(open(f,encoding='utf-8').read()) for f in glob.glob('skills/k3-bom-grid-entry/scripts/*.py')];print('ok')"
```

6. 涉及坐标变更时，同步更新 `SKILL.md` 的「步骤 3 标定」表格与 `README.md` 的坐标表。

## PR 流程

1. Fork → 新建分支 `fix/xxx` 或 `feat/xxx`；
2. 改动 + 自测（在测试账套上跑通一次真实录入）；
3. 在 `CHANGELOG.md` 的 `Unreleased` 段落加一条；
4. 提 PR，描述里写清：**K3 版本、Windows 版本、分辨率、DPI 缩放、是否分屏、每页行数**。

## 提交信息格式

```
<type>: <简述>

type ∈ feat | fix | docs | refactor | chore
```

例：`fix: 用量列坐标按 rect.left 取最小者，避免命中备注列`

## 行为准则

讨论对事不对人。涉及真实生产数据的截图、日志请先脱敏。
