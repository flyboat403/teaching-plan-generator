"""Reusable DOCX teaching plan generator.

Usage:
    1. pip install python-docx
    2. Organize teaching plan fields into JSON (see structure below)
    3. Write JSON to a temporary file
    4. python generate_docx.py <input.json> [output.docx]
    5. Delete the temporary JSON file

Input JSON structure:
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
        ]},
        ...
    ],
    "output_path": "output.docx"  (optional, can be 2nd argv)
}

content types:
  - "text" (default): string or array of {text, bold, indent} dicts
  - "tables": array of {header, rows}
"""
import json
import sys
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── styles ──────────────────────────────────────────────────
FONT_TITLE   = '黑体'
FONT_HEADING = '黑体'
FONT_BODY    = '宋体'
FONT_TABLE   = '宋体'
FONT_TABLE_H = '黑体'
SIZE_TITLE   = Pt(18)
SIZE_HEADING = Pt(16)
SIZE_BODY    = Pt(12)
SIZE_TABLE   = Pt(10.5)


def _set_run_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run.font.size = size
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def _add_styled_para(doc, text, font=FONT_BODY, size=SIZE_BODY, bold=False,
                     alignment=None, indent=False, spacing=1.5):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = spacing
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    if alignment is not None:
        p.alignment = alignment
    run = p.add_run(text)
    _set_run_font(run, font, size, bold)
    return p


def _add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    _set_run_font(run, FONT_TITLE, SIZE_TITLE, bold=True)


def _add_heading_styled(doc, text):
    h = doc.add_heading(text, level=1)
    for run in h.runs:
        _set_run_font(run, FONT_HEADING, SIZE_HEADING, bold=True)


def _rebuild_table_row(row, values, bold=False):
    """Replace cell text in a table row."""
    if not isinstance(values, list):
        raise ValueError(f'表格行必须是数组，实得 {type(values).__name__}: {values!r}')
    if len(values) != len(row.cells):
        raise ValueError(f'表格行列数不匹配: 该行应有 {len(row.cells)} 列, 实得 {len(values)} 列: {values}')
    font = FONT_TABLE_H if bold else FONT_TABLE
    for i, val in enumerate(values):
        cell = row.cells[i]
        # 真正清空单元格：python-docx 的 cell.text="" 只设置第一个段落，
        # 旧 run 会残留导致字体/内容串线，必须显式删除
        for extra_p in cell.paragraphs[1:]:
            extra_p._element.getparent().remove(extra_p._element)
        p = cell.paragraphs[0]
        for old_run in list(p.runs):
            old_run._element.getparent().remove(old_run._element)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(str(val))
        _set_run_font(run, font, SIZE_TABLE, bold=bold)


def _add_table(doc, header, rows):
    """Add a formatted table with header row."""
    if not header:
        raise ValueError('表格 header 不能为空')
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    _rebuild_table_row(table.rows[0], header, bold=True)
    for r_idx, row in enumerate(rows):
        _rebuild_table_row(table.rows[r_idx + 1], row)


def _render_content(doc, content):
    """Render a single content item: plain text or nested list."""
    if isinstance(content, str):
        _add_styled_para(doc, content, indent=True)
    elif isinstance(content, list):
        for item in content:
            _render_content(doc, item)
    elif isinstance(content, dict):
        text = content.get('text') or ''
        bold = bool(content.get('bold', False))
        indent = bool(content.get('indent', True))
        _add_styled_para(doc, text, bold=bold, indent=indent)
    else:
        raise ValueError('不支持的 content 类型: %s (值: %r)' % (type(content).__name__, content))


def generate(data, output_path):
    """Generate DOCX from structured data."""
    doc = Document()

    # Page setup
    sec = doc.sections[0]
    sec.page_width  = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin    = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(3.0)

    fields = data.get('fields', [])

    for i, field in enumerate(fields):
        if not isinstance(field, dict):
            raise ValueError(f'fields[{i}] 必须是对象，实得 {type(field).__name__}')
        heading = field.get('heading', '')
        if not isinstance(heading, str):
            raise ValueError(f'fields[{i}].heading 必须是字符串，实得 {type(heading).__name__}：{heading!r}')
        content = field.get('content', '')
        ctype   = field.get('type', 'text')

        if i == 0:
            title = data.get('title') or heading
            if not isinstance(title, str):
                raise ValueError(f'title 必须是字符串，实得 {type(title).__name__}：{title!r}')
            _add_title(doc, title)

            _add_heading_styled(doc, heading)
        else:
            _add_heading_styled(doc, heading)

        if ctype not in ('text', 'tables'):
            raise ValueError(f'不支持的 type: {ctype!r}（仅支持 text / tables）')
        if ctype == 'tables':
            for tbl in content:
                if not isinstance(tbl, dict) or not isinstance(tbl.get('rows'), list):
                    raise ValueError('tables 的每个元素必须是包含 rows 数组的对象: %r' % (tbl,))
                _add_table(doc, tbl.get('header', []), tbl.get('rows', []))
            # Add blank paragraph after tables
            _add_styled_para(doc, '')
        else:
            _render_content(doc, content)

    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    size_kb = round(os.path.getsize(output_path) / 1024, 1)
    print(f'DOCX saved: {output_path} ({size_kb} KB)')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python generate_docx.py <input.json> [output.docx]')
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path.replace('.json', '.docx')

    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print('错误: 找不到输入文件 ' + input_path); sys.exit(1)
    except json.JSONDecodeError as e:
        print('错误: JSON 非法 (%s, 第%d行第%d列)，请核对 Output Format Spec' % (e.msg, e.lineno, e.colno)); sys.exit(1)

    if not isinstance(data, dict) or not isinstance(data.get('fields'), list) or not data['fields']:
        print('错误: fields 必须是包含至少一个元素的对象数组'); sys.exit(1)

    if 'output_path' in data and len(sys.argv) <= 2:
        output_path = data['output_path']
    if not isinstance(output_path, str) or not output_path:
        print('错误: output_path 必须是字符串'); sys.exit(1)

    try:
        generate(data, output_path)
    except ValueError as e:
        print('错误: ' + str(e)); sys.exit(1)
