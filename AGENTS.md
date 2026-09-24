# AGENTS.md — teaching-plan-generator 仓库导览

本仓库是一个 Agent 技能包（非独立应用）。**`SKILL.md` 是唯一事实源**——生成教案、加载模板、DOCX 调用规则、质量自检等全部以 SKILL.md 为准，本文件不复述其内容。

## 文件结构

| 文件 | 用途 |
|------|------|
| `SKILL.md` | 技能定义：工作流、字段规则、学情决策树、模板加载、质量自检 |
| `references/uniform-template.md` | 统一学情模板（4列表格）——学情无分层时加载 |
| `references/layered-template.md` | 分层学情模板（5列表格，含分层标注）——学情分层时加载 |
| `references/docx-style-spec.md` | DOCX 样式规范（页面/字体/表格/编号）——生成 DOCX 时遵循 |
| `scripts/generate_docx.py` | DOCX 生成脚本——只通过 JSON 调用，NEVER 重写 |

## 仓库约定

- 技能为中文技能：交互与输出一律使用中文
- 临时 JSON 文件在 DOCX 生成成功后立即删除
- 模板占位符使用 `{{Mustache}}` 语法，`{{#each}}...{{/each}}` 为可重复区块
- 输出文件命名：`[课程名称]_[课题名称]_教学设计.md` / `.docx`

## 依赖速查

| 包 | 用途 |
|----|------|
| `python-docx` | DOCX 生成 + DOCX 参考材料读取（必需） |
| `pypdf` / `PyMuPDF` | PDF 参考材料（首选/备选，pypdf 轻量、失败时再装 PyMuPDF） |
| `pytesseract` + `Pillow` | 图片参考材料 OCR（需 Tesseract 引擎） |

安装命令与解析回退链详见 SKILL.md § Dependencies 与 § Reference Material Processing。
