# AGENTS.md — teaching-plan-generator 仓库导览

> **非加载面**：Agent 执行本技能时**不要**读取本文件；只读 `SKILL.md` 及其触发的 `references/`、`scripts/`。本文件仅供人类维护仓库。

**`SKILL.md` 是唯一事实源。**

## 文件结构

| 文件 | 用途 |
|------|------|
| `SKILL.md` | 技能定义：工作流、决策树、NEVER、质量自检 |
| `references/uniform-template.md` | 统一学情模板——Step 2 加载 |
| `references/layered-template.md` | 分层学情模板——Step 2 加载 |
| `references/field-rules.md` | 逐字段生成规则——Step 3 加载 |
| `references/docx-style-spec.md` | DOCX 样式——Step 5 对照 §8 |
| `references/docx-json-contract.md` | DOCX JSON 契约——Step 4 加载 |
| `scripts/generate_docx.py` | DOCX 脚本——只经 JSON 调用，NEVER 重写 |
| `examples/` | 示例产物（可选，非加载面） |

## 仓库约定

- 中文交互与输出；临时 JSON 成功后即删；输出命名 `[课程]_[课题]_教学设计.md/.docx`
- 依赖安装与回退链见 `SKILL.md` § Reference Material Processing 与 Step 4
