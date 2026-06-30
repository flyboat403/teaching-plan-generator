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
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


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
    font = FONT_TABLE_H if bold else FONT_TABLE
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(str(val))
        _set_run_font(run, font, SIZE_TABLE, bold=bold)


def _add_table(doc, header, rows):
    """Add a formatted table with header row."""
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
        text = content.get('text', '')
        bold = content.get('bold', False)
        indent = content.get('indent', True)
        _add_styled_para(doc, text, bold=bold, indent=indent)


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
        heading = field.get('heading', '')
        content = field.get('content', '')
        ctype   = field.get('type', 'text')

        if i == 0:
            # First field heading is the document title
            title_text = heading.replace('一、', '').replace('课程名称', '').strip()
            # Use title from data if provided
            _add_title(doc, data.get('title', heading))
            _add_heading_styled(doc, heading)
        else:
            _add_heading_styled(doc, heading)

        if ctype == 'tables':
            for tbl in content:
                _add_table(doc, tbl.get('header', []), tbl.get('rows', []))
            # Add blank paragraph after tables
            _add_styled_para(doc, '')
        else:
            _render_content(doc, content)

    doc.save(output_path)
    size_kb = round(os.path.getsize(output_path) / 1024, 1)
    print(f'DOCX saved: {output_path} ({size_kb} KB)')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python generate_docx.py <input.json> [output.docx]')
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path.replace('.json', '.docx')

    with open(input_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if 'output_path' in data and len(sys.argv) <= 2:
        output_path = data['output_path']

    generate(data, output_path)
