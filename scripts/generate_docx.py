"""Reusable DOCX teaching plan generator (style per references/docx-style-spec.md).

Usage:
    1. pip install python-docx
    2. Organize teaching plan into JSON (see structure below)
    3. Write JSON to a temporary file
    4. python3 generate_docx.py <input.json> [output.docx]
    5. Delete the temporary JSON file

Input JSON structure:
{
    "title": "《课程名称》「课题名称」教学设计",
    "basic_info": [
        ["课程名称", "办公软件应用"],
        ["课题名称", "文字环绕"],
        ...  # 9 label-value rows -> unnumbered info table (4.0 / 11.0 cm)
    ],
    "sections": [
        {"heading": "一、学情分析", "content": "..."},
        {"heading": "二、教学目标", "content": [
            {"text": "知识目标", "bold": true},
            {"text": "1. ...", "indent": true}
        ]},
        {"heading": "九、教学流程", "type": "tables", "content": [
            {"header": ["时间/环节", "教师活动", "学生活动", "设计意图"],
             "widths": [2.0, 6.4, 3.4, 3.2],
             "rows": [
                {"band": "【第一学时 0—40分钟】1.1 课题"},
                ["0-2min 导入锚定", "...", "...", "..."],
                {"merged": "第一次缓冲（10—13分钟）：...，不讲新知识"},
                {"cells": ["合计", "", "100分", "", ""],
                 "merges": [[0, 1], [3, 4]], "total": true}
             ],
             "page_break_before": false}
        ]}
    ],
    "output_path": "output.docx"  (optional, can be 2nd argv)
}

content types (the 'type' field is OPTIONAL - inferred from content shape):
  - "text" (default): string or array of {text, bold, indent} dicts
  - "tables": array of {header, rows, widths?, page_break_before?} (or a single such dict)

table row types (entries in "rows"):
  - ["cell", ...]                data row (宋体 12pt, left, no shading)
  - {"band": "text"}             merged full-width band (黑体 12pt bold, EDEDED)
  - {"merged": "text"}           merged full-width data-styled row (buffers / 终局复盘)
  - {"cells": [...], "merges": [[i,j],...], "total": true}
                                 total row (黑体 bold centered, E7E7E7)

Style source: references/docx-style-spec.md (page/margins/fonts/tables/footer/numbering).
"""
import json
import sys
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


FONT_TITLE = '黑体'
FONT_HEADING = '黑体'
FONT_BODY = '宋体'
FONT_BOARD = '宋体'

SIZE_TITLE = Pt(18)
SIZE_HEADING = Pt(16)
SIZE_BODY = Pt(12)
SIZE_BOARD = Pt(10.5)
SIZE_TABLE = Pt(12)
SIZE_INFO = Pt(10.5)
SIZE_FOOTER = Pt(9)

SHD_HEADER = 'D9D9D9'
SHD_BAND = 'EDEDED'
SHD_TOTAL = 'E7E7E7'
SHD_INFO_LABEL = 'F2F2F2'

USABLE_WIDTH_CM = 15.0
WIDTHS_FLOW_4 = [2.0, 6.4, 3.4, 3.2]
WIDTHS_FLOW_5 = [2.0, 5.6, 3.4, 2.6, 1.4]
WIDTHS_EVAL_5 = [2.2, 3.4, 5.4, 2.2, 1.8]
WIDTHS_INFO = [4.0, 11.0]

_BOARD_PREFIXES = ('├', '└', '│', '┌', '┐', '┘', '┤', '┬', '┴', '┼', '═', '╠', '╣', '║')


def _set_run_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run.font.size = size
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)


def _set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.insert_element_before(
            shd, 'w:noWrap', 'w:tcMar', 'w:textDirection',
            'w:tcFitText', 'w:vAlign', 'w:hideMark')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)


def _set_cell_valign_center(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        tcPr.insert_element_before(vAlign, 'w:hideMark')
    vAlign.set(qn('w:val'), 'center')


def _set_row_header_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = trPr.find(qn('w:tblHeader'))
    if tblHeader is None:
        tblHeader = OxmlElement('w:tblHeader')
        trPr.insert_element_before(tblHeader, 'w:cantSplit', 'w:trHeight')


def _set_table_fixed(table):
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')


def _apply_column_widths(table, widths_cm):
    for i, w in enumerate(widths_cm):
        table.columns[i].width = Cm(w)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)


def _clear_cell(cell):
    """Remove surplus paragraphs and runs so exactly one empty paragraph remains."""
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    for child in list(p._p):
        if child.tag != qn('w:pPr'):
            p._p.remove(child)
    return p


def _fill_cell(cell, text, font, size, bold, align,
               before, after, line, valign=False, shading=None):
    p = _clear_cell(cell)
    p.paragraph_format.space_before = before
    p.paragraph_format.space_after = after
    p.paragraph_format.line_spacing = line
    if align is not None:
        p.alignment = align
    run = p.add_run(str(text))
    _set_run_font(run, font, size, bold)
    if shading:
        _set_cell_shading(cell, shading)
    if valign:
        _set_cell_valign_center(cell)


def _looks_like_board(text):
    t = (text or '').lstrip(' \t　')
    return bool(t) and t.startswith(_BOARD_PREFIXES)


def _setup_styles(doc):
    style = doc.styles['Heading 1']
    style.font.name = FONT_HEADING
    style.font.size = SIZE_HEADING
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(10)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), FONT_HEADING)
    rFonts.set(qn('w:ascii'), FONT_HEADING)
    rFonts.set(qn('w:hAnsi'), FONT_HEADING)
    c = rPr.find(qn('w:color'))
    if c is None:
        c = OxmlElement('w:color')
        rPr.append(c)
    for attr in ('w:themeColor', 'w:themeShade', 'w:themeTint'):
        if c.get(qn(attr)) is not None:
            del c.attrib[qn(attr)]
    c.set(qn('w:val'), '000000')

    normal = doc.styles['Normal']
    normal.font.name = FONT_BODY
    nPr = normal.element.get_or_add_rPr()
    nFonts = nPr.find(qn('w:rFonts'))
    if nFonts is None:
        nFonts = OxmlElement('w:rFonts')
        nPr.insert(0, nFonts)
    nFonts.set(qn('w:eastAsia'), FONT_BODY)
    nFonts.set(qn('w:ascii'), FONT_BODY)
    nFonts.set(qn('w:hAnsi'), FONT_BODY)


def _setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(3.0)
    sec.header_distance = Cm(1.27)
    sec.footer_distance = Cm(1.27)
    sec.header.is_linked_to_previous = True


def _add_page_field(paragraph, instr):
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:eastAsia'), FONT_BODY)
    rf.set(qn('w:ascii'), FONT_BODY)
    rf.set(qn('w:hAnsi'), FONT_BODY)
    rPr.append(rf)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')
    rPr.append(sz)
    run.append(rPr)
    fld.append(run)
    paragraph._p.append(fld)


def _setup_footer(section):
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for child in list(p._p):
        if child.tag != qn('w:pPr'):
            p._p.remove(child)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0

    def add_text(text):
        run = p.add_run(text)
        _set_run_font(run, FONT_BODY, SIZE_FOOTER, False)

    add_text('第 ')
    _add_page_field(p, ' PAGE ')
    add_text(' 页  共 ')
    _add_page_field(p, ' NUMPAGES ')
    add_text(' 页')


def _add_styled_para(doc, text, font=FONT_BODY, size=SIZE_BODY, bold=False,
                     alignment=None, indent=False, spacing=1.5,
                     before=None, after=None):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = spacing
    if before is not None:
        p.paragraph_format.space_before = before
    if after is not None:
        p.paragraph_format.space_after = after
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
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    _set_run_font(run, FONT_TITLE, SIZE_TITLE, bold=True)


def _add_heading_styled(doc, text):
    h = doc.add_heading(text, level=1)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.line_spacing = 1.0
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        _set_run_font(run, FONT_HEADING, SIZE_HEADING, bold=True)
        run.font.color.rgb = RGBColor(0, 0, 0)
        rPr = run._element.get_or_add_rPr()
        c = rPr.find(qn('w:color'))
        if c is None:
            c = OxmlElement('w:color')
            rPr.append(c)
        for attr in ('w:themeColor', 'w:themeShade', 'w:themeTint'):
            if c.get(qn(attr)) is not None:
                del c.attrib[qn(attr)]
        c.set(qn('w:val'), '000000')
    return h


def _render_content(doc, content):
    if isinstance(content, str):
        if content == '':
            _add_styled_para(doc, '', indent=False, spacing=1.5)
        else:
            _add_styled_para(doc, content, indent=True)
        return
    if isinstance(content, list):
        for item in content:
            _render_content(doc, item)
        return
    if isinstance(content, dict):
        if 'text' not in content:
            raise ValueError('text 段落缺少 text 字段: %r' % (content,))
        text = content['text'] or ''
        bold = bool(content.get('bold', False))
        style = content.get('style')
        if 'indent' in content:
            indent = bool(content['indent'])
        else:
            indent = not bold
        if style == 'board' or (style is None and _looks_like_board(text)):
            _add_styled_para(doc, text, font=FONT_BOARD, size=SIZE_BOARD,
                             bold=False, indent=False, spacing=1.0,
                             before=Pt(0), after=Pt(0))
        elif bold and not indent:
            _add_styled_para(doc, text, bold=True, indent=False,
                             spacing=1.5, before=Pt(8), after=Pt(2))
        else:
            _add_styled_para(doc, text, bold=bold, indent=indent)
        return
    raise ValueError('不支持的 content 类型: %s (值: %r)' % (type(content).__name__, content))


def _is_table_spec(obj):
    return isinstance(obj, dict) and (
        isinstance(obj.get('rows'), list) or isinstance(obj.get('header'), list))


def _detect_content_type(content):
    if isinstance(content, str):
        return 'text'
    if isinstance(content, dict) and _is_table_spec(content):
        return 'tables'
    if isinstance(content, list):
        if not content:
            return 'text'
        if any(_is_table_spec(item) for item in content):
            return 'tables'
        return 'text'
    raise ValueError('不支持的 content 类型: %s (值: %r)' % (type(content).__name__, content))


def _resolve_widths(heading, ncols, explicit):
    if explicit is not None:
        if not isinstance(explicit, list) or len(explicit) != ncols:
            raise ValueError(
                'widths 必须是长度为 %d 的数组，实得 %r' % (ncols, explicit))
        try:
            widths = [float(w) for w in explicit]
        except (TypeError, ValueError):
            raise ValueError('widths 必须为数字数组: %r' % (explicit,))
        if abs(sum(widths) - USABLE_WIDTH_CM) > 0.05:
            raise ValueError(
                'widths 合计必须为 15.0 cm（可用正文宽），实得 %.2f: %r'
                % (sum(widths), widths))
        return widths

    if '流程' in heading:
        if ncols == 5:
            return list(WIDTHS_FLOW_5)
        if ncols == 4:
            return list(WIDTHS_FLOW_4)
    if '评价' in heading and ncols == 5:
        return list(WIDTHS_EVAL_5)

    unit = round(USABLE_WIDTH_CM / ncols, 2)
    widths = [unit] * ncols
    widths[-1] = round(USABLE_WIDTH_CM - unit * (ncols - 1), 2)
    return widths


def _style_header_cell(cell, text):
    _fill_cell(cell, text, FONT_HEADING, SIZE_TABLE, True,
               WD_ALIGN_PARAGRAPH.CENTER, Pt(2), Pt(2), 1.5,
               valign=True, shading=SHD_HEADER)


def _style_data_cell(cell, text):
    _fill_cell(cell, text, FONT_BODY, SIZE_TABLE, False,
               WD_ALIGN_PARAGRAPH.LEFT, Pt(2), Pt(2), 1.5,
               valign=False, shading=None)


def _style_band_cell(cell, text):
    _fill_cell(cell, text, FONT_HEADING, SIZE_TABLE, True,
               WD_ALIGN_PARAGRAPH.LEFT, Pt(2), Pt(2), 1.5,
               valign=True, shading=SHD_BAND)


def _style_total_cell(cell, text):
    _fill_cell(cell, text, FONT_HEADING, SIZE_TABLE, True,
               WD_ALIGN_PARAGRAPH.CENTER, Pt(2), Pt(2), 1.5,
               valign=True, shading=SHD_TOTAL)


def _merge_row_span(row, ncols):
    """Merge columns 0..ncols-1 into one cell; return the merged cell."""
    merged = row.cells[0]
    for i in range(1, ncols):
        merged = merged.merge(row.cells[i])
    return merged


def _render_row(table, row_idx, row_data, ncols, heading):
    row = table.rows[row_idx]
    if isinstance(row_data, dict):
        if 'band' in row_data:
            cell = _merge_row_span(row, ncols)
            _style_band_cell(cell, row_data['band'])
            return
        if 'merged' in row_data:
            cell = _merge_row_span(row, ncols)
            _fill_cell(cell, row_data['merged'], FONT_BODY, SIZE_TABLE, False,
                       WD_ALIGN_PARAGRAPH.LEFT, Pt(2), Pt(2), 1.5)
            return
        if 'cells' in row_data:
            cells = row_data['cells']
            if not isinstance(cells, list) or len(cells) != ncols:
                raise ValueError(
                    '合计行 cells 必须是长度为 %d 的数组，实得 %r' % (ncols, cells))
            merges = row_data.get('merges') or []
            if not isinstance(merges, list):
                raise ValueError('merges 必须是数组: %r' % (merges,))
            for rng in merges:
                if (not isinstance(rng, (list, tuple)) or len(rng) != 2
                        or not all(isinstance(x, int) for x in rng)):
                    raise ValueError('merges 每项须为 [start, end] 整数对: %r' % (rng,))
                start, end = rng
                if start < 0 or end >= ncols or start >= end:
                    raise ValueError('merges 范围非法: %r（列数 %d）' % (rng, ncols))
                row.cells[start].merge(row.cells[end])
            seen = set()
            styler = _style_total_cell if row_data.get('total') else _style_data_cell
            for i, val in enumerate(cells):
                tc = row.cells[i]._tc
                if id(tc) in seen:
                    continue
                seen.add(id(tc))
                styler(row.cells[i], val)
            return
        raise ValueError(
            '无法识别的行对象（需含 band / merged / cells 之一）: %r' % (row_data,))

    if isinstance(row_data, list):
        if len(row_data) != ncols:
            raise ValueError(
                '表格行列数不匹配（%s）: 该行应有 %d 列, 实得 %d 列: %r'
                % (heading, ncols, len(row_data), row_data))
        for i, val in enumerate(row_data):
            _style_data_cell(row.cells[i], val)
        return

    raise ValueError('表格行必须是数组或行对象，实得 %s: %r'
                     % (type(row_data).__name__, row_data))


def _add_spec_table(doc, spec, heading):
    if not isinstance(spec, dict):
        raise ValueError('表格规格必须是对象: %r' % (spec,))
    header = spec.get('header')
    rows = spec.get('rows', [])
    if not isinstance(header, list) or not header:
        raise ValueError('表格 header 不能为空（%s）' % heading)
    if not isinstance(rows, list):
        raise ValueError('表格 rows 必须是数组（%s）' % heading)

    ncols = len(header)
    widths = _resolve_widths(heading, ncols, spec.get('widths'))

    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    _set_table_fixed(table)
    _apply_column_widths(table, widths)

    hrow = table.rows[0]
    for i, text in enumerate(header):
        _style_header_cell(hrow.cells[i], text)
    _set_row_header_repeat(hrow)

    for r_idx, row_data in enumerate(rows):
        _render_row(table, r_idx + 1, row_data, ncols, heading)

    if spec.get('page_break_before'):
        first = table.rows[0].cells[0].paragraphs[0]
        first.paragraph_format.page_break_before = True

    _add_styled_para(doc, '', indent=False, spacing=1.0, before=Pt(0), after=Pt(0))
    return table


def _add_basic_info_table(doc, basic_info):
    if not isinstance(basic_info, list) or not basic_info:
        raise ValueError('basic_info 必须是非空数组，元素为 [标签, 值] 二元组')
    for i, row in enumerate(basic_info):
        if not isinstance(row, (list, tuple)) or len(row) != 2:
            raise ValueError('basic_info[%d] 必须是 [标签, 值] 二元组，实得 %r' % (i, row))

    table = doc.add_table(rows=len(basic_info), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    _set_table_fixed(table)
    _apply_column_widths(table, WIDTHS_INFO)

    for i, (label, value) in enumerate(basic_info):
        row = table.rows[i]
        _fill_cell(row.cells[0], label, FONT_HEADING, SIZE_INFO, True,
                   WD_ALIGN_PARAGRAPH.CENTER, Pt(1), Pt(1), 1.0,
                   valign=True, shading=SHD_INFO_LABEL)
        _fill_cell(row.cells[1], value, FONT_BODY, SIZE_INFO, False,
                   WD_ALIGN_PARAGRAPH.LEFT, Pt(1), Pt(1), 1.0,
                   valign=True, shading=None)

    _add_styled_para(doc, '', indent=False, spacing=1.0, before=Pt(0), after=Pt(0))
    return table


def generate(data, output_path):
    """Generate DOCX from structured data (basic_info + 12 sections)."""
    doc = Document()
    _setup_styles(doc)
    _setup_page(doc)
    _setup_footer(doc.sections[0])

    title = data.get('title')
    if title is not None:
        if not isinstance(title, str):
            raise ValueError('title 必须是字符串，实得 %s' % type(title).__name__)
        if title:
            _add_title(doc, title)

    basic_info = data.get('basic_info')
    if basic_info is not None:
        _add_basic_info_table(doc, basic_info)

    sections = data.get('sections', [])
    if not isinstance(sections, list) or not sections:
        raise ValueError('sections 必须是包含至少一个元素的数组')

    for i, section in enumerate(sections):
        if not isinstance(section, dict):
            raise ValueError('sections[%d] 必须是对象，实得 %s'
                             % (i, type(section).__name__))
        heading = section.get('heading', '')
        if not isinstance(heading, str) or not heading:
            raise ValueError('sections[%d].heading 必须是非空字符串' % i)
        content = section.get('content', '')
        ctype = section.get('type', None)

        _add_heading_styled(doc, heading)

        detected = _detect_content_type(content)
        if ctype is None:
            ctype = detected
        elif ctype not in ('text', 'tables'):
            raise ValueError('不支持的 type: %r（仅支持 text / tables）' % (ctype,))
        elif ctype != detected:
            raise ValueError(
                'sections[%d] (%s) 的 type=%r 与内容实际形状 %r 不一致，请检查 JSON'
                % (i, heading, ctype, detected))

        if ctype == 'tables':
            specs = content if isinstance(content, list) else [content]
            for spec in specs:
                _add_spec_table(doc, spec, heading)
        else:
            _render_content(doc, content)

    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    size_kb = round(os.path.getsize(output_path) / 1024, 1)
    print('DOCX saved: %s (%s KB)' % (output_path, size_kb))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 generate_docx.py <input.json> [output.docx]')
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path.replace('.json', '.docx')

    try:
        with open(input_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except FileNotFoundError:
        print('错误: 找不到输入文件 ' + input_path)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print('错误: JSON 非法 (%s, 第%d行第%d列)，请核对 references/docx-json-contract.md'
              % (e.msg, e.lineno, e.colno))
        sys.exit(1)

    if not isinstance(data, dict):
        print('错误: JSON 根必须是对象')
        sys.exit(1)
    if 'fields' in data and 'sections' not in data:
        print('错误: 旧版 fields 结构已废弃，请改用 basic_info + sections（见 references/docx-json-contract.md）')
        sys.exit(1)
    if not isinstance(data.get('sections'), list) or not data['sections']:
        print('错误: sections 必须是包含至少一个元素的数组')
        sys.exit(1)

    if 'output_path' in data and len(sys.argv) <= 2:
        output_path = data['output_path']
    if not isinstance(output_path, str) or not output_path:
        print('错误: output_path 必须是字符串')
        sys.exit(1)

    try:
        generate(data, output_path)
    except ValueError as e:
        print('错误: ' + str(e))
        sys.exit(1)
