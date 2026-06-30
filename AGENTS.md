# AGENTS.md — teaching-plan-generator

An OpenCode **skill** (not a standalone app) that generates 40-minute three-phase teaching plans (教案) for vocational education, outputting MD + DOCX.

## Entry point

`SKILL.md` is the single source of truth. **Read it fully before generating any teaching plan.** It contains all rules, field definitions, template selection logic, quality checklist, and output format specs.

## File structure

| File | Purpose |
|------|---------|
| `SKILL.md` | Skill definition — workflow, rules, field generation, quality checklist |
| `scripts/generate_docx.py` | Reusable DOCX generator — **do not rewrite**, call via JSON input |
| `references/uniform-template.md` | Template for uniform-proficiency classes (4-column tables, no 分层标注) |
| `references/layered-template.md` | Template for differentiated classes (5-column tables with 分层标注, dual objectives) |

## Template selection (decision tree)

Parse the teacher's student-proficiency description (学情描述):

- Contains keywords like 分层/两极分化/基础弱/差异大/一部分...另一部分 → **layered** template
  - Must ask for split ratio (e.g., "60% 基础 / 40% 进阶")
- Uniform description or unclear → **uniform** template
- If cannot decide → offer A/B choice to teacher

## DOCX generation

**Never rewrite `scripts/generate_docx.py`.** Always call it with JSON:

```bash
pip install python-docx
python scripts/generate_docx.py <input.json> [output.docx]
```

JSON schema for the script:

```json
{
  "title": "《课程名称》「课题名称」教学设计",
  "fields": [
    {"heading": "一、课程名称", "content": "办公软件应用"},
    {"heading": "二、课题名称", "content": "图片环绕方式"},
    ...
    {"heading": "十、教学目标", "content": [
      {"text": "基础目标（全体学生达到）", "bold": true},
      {"text": "知识目标描述", "indent": true}
    ]},
    {"heading": "十八、教学流程", "type": "tables", "content": [
      {"header": ["时间/环节","教师活动","学生活动","设计意图"],
       "rows": [["0-2min 导入锚定","...","...","..."]]}
    ]}
  ]
}
```

`content` types: `"text"` (default, string or array of `{text, bold, indent}` dicts) or `"tables"` (array of `{header, rows}`).

## Core constraints (from SKILL.md)

- **Time formula**: 10 + 3 + 10 + 3 + 10 + 4 = 40 min (verify each generation)
- **Module 1** (0-10min): basics only, no extension
- **Module 2** (13-23min): difficulties/error-prone only, no repeat of module 1
- **Module 3** (26-36min): practice/case study only, no new knowledge
- **Buffers** (10-13min, 23-26min): no new knowledge, pick one of three options
- **Opening**: must use real industry scenario or enterprise problem
- **21 fields**: 1-9 basic info, 10-21 pedagogical fields (generate in order)
- **Teaching objectives**: use observable verbs (说出/操作/诊断), never 了解/知道/掌握
- **Layered version**: split objectives into 基础/进阶, add 分层标注 column, dual-standard evaluation

## Dependencies

| Package | Use case | Command |
|---------|----------|---------|
| `python-docx` | DOCX output + reading DOCX references | `pip install python-docx` |
| `pdfplumber` | PDF reference extraction (preferred) | `pip install pdfplumber` |
| `PyMuPDF` | PDF reference extraction (fallback) | `pip install PyMuPDF` |
| `pytesseract` | Image OCR for reference materials | `pip install pytesseract` + Tesseract engine |
| `Pillow` | Image preprocessing | `pip install Pillow` |

If teacher provides reference materials, classify as A (content doc), B (existing plan), or C (enterprise manual) before extraction.

## Quality self-check (after every generation)

Verify all 10 items from SKILL.md § Quality Checklist. Key ones agents commonly miss:

- [ ] Time sum = 40 min exactly
- [ ] Each module has exactly 1 core knowledge point
- [ ] Buffers contain **no new knowledge**
- [ ] Objectives use observable verbs (no 了解/知道/掌握)
- [ ] Opening uses real industry scenario
- [ ] Layered version: split objectives, column, dual standards all present

## Style conventions (for this repo)

- Chinese-language skill (all content, user interaction, and output in Chinese)
- Store temporary JSON files, remove after DOCX generation
- Template placeholders use `{{Mustache}}` syntax; `{{#each}}...{{/each}}` indicates repeatable blocks
- File naming: `[课程名称]_[课题名称]_教学设计.md` / `.docx`

## What this repo is NOT

- Not a standalone application — no package.json, no build step, no framework
- Not a git repo — no CI/CD, no tests
- The AGENTS.md itself — read it on session start for onboarding
