"""User Word styles and reusable SCI-style layout helpers (python-docx).

No PDF converter and no experiment-specific paths or data.
"""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

TEMPLATE = Path(__file__).resolve().parents[1] / 'assets' / 'word-styles-template.docx'
ALIASES = {
    'body': ('正文1',),
    'heading1': ('标题1',),
    'heading2': ('标题 2', 'Heading 2', 'heading 2'),
    'heading3': ('标题 3', 'Heading 3', 'heading 3'),
    'image': ('TOC 1', 'toc 1'),
}


def style_for(doc, role):
    for name in ALIASES[role]:
        try:
            return doc.styles[name]
        except KeyError:
            continue
    raise ValueError('Required user style missing: ' + role + '; inspect the supplied template')


def _black(run):
    run.font.color.rgb = RGBColor(0, 0, 0)
    return run


def _no_indent(p):
    p.paragraph_format.first_line_indent = Cm(0)
    p._p.get_or_add_pPr().get_or_add_ind().set(qn('w:firstLineChars'), '0')


def _no_grid(p):
    pp = p._p.get_or_add_pPr()
    for old in pp.findall(qn('w:snapToGrid')):
        pp.remove(old)
    e = OxmlElement('w:snapToGrid')
    e.set(qn('w:val'), '0')
    pp.append(e)


def _available_width(doc):
    sec = doc.sections[-1]
    return sec.page_width - sec.left_margin - sec.right_margin


def set_page_numbers(doc):
    """Hide numbering on document's first page; use PAGE field afterwards."""
    for i, sec in enumerate(doc.sections):
        sec.different_first_page_header_footer = (i == 0)
        if i:
            sec.footer.is_linked_to_previous = True
            continue
        sec.footer.is_linked_to_previous = False
        sec.first_page_footer.is_linked_to_previous = False
        for footer in (sec.footer, sec.first_page_footer):
            for child in list(footer._element):
                footer._element.remove(child)
            footer.add_paragraph('', style=style_for(doc, 'body'))
        p = sec.footer.paragraphs[0]
        _no_indent(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')
        p._p.append(field)
    settings = doc.settings.element
    for old in settings.findall(qn('w:updateFields')):
        settings.remove(old)
    update = OxmlElement('w:updateFields')
    update.set(qn('w:val'), 'true')
    settings.append(update)


def new_document(template=None):
    """Create a new report from styles; never use to edit an existing report."""
    doc = Document(str(Path(template) if template else TEMPLATE))
    for child in list(doc.element.body):
        if child.tag != qn('w:sectPr'):
            doc.element.body.remove(child)
    for role in ALIASES:
        style_for(doc, role)
    # New reports do not add their own header/cover metadata.
    for sec in doc.sections:
        for header in (sec.header, sec.first_page_header, sec.even_page_header):
            for child in list(header._element):
                header._element.remove(child)
            header.add_paragraph('', style=style_for(doc, 'body'))
    set_page_numbers(doc)
    return doc


def add_body(doc, text):
    p = doc.add_paragraph(style=style_for(doc, 'body'))
    _black(p.add_run(text))
    return p


def add_heading(doc, text, level=1):
    if level not in (1, 2, 3):
        raise ValueError('Use user heading levels 1, 2 or 3')
    first = not doc.paragraphs
    p = doc.add_paragraph(style=style_for(doc, 'heading' + str(level)))
    _black(p.add_run(text))
    p.paragraph_format.keep_with_next = True
    if first:
        p.paragraph_format.page_break_before = False
    return p


def _next_number(doc, pattern):
    found = [int(m.group(1)) for p in doc.paragraphs
             if (m := re.match(pattern, p.text))]
    return max(found, default=0) + 1


def _table_captions(tb):
    description = tb._tbl.getprevious()
    label = description.getprevious() if description is not None else None
    if (label is None or label.tag != qn('w:p') or description.tag != qn('w:p')):
        return []
    a, b = Paragraph(label, tb._parent), Paragraph(description, tb._parent)
    return [a, b] if re.fullmatch(r'Table\s+\d+', a.text.strip()) else []


def keep_table_together(tb, captions=None):
    """Join caption pair and every row; real layout still needs a page check."""
    for p in (_table_captions(tb) if captions is None else captions):
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.keep_together = True
    for i, row in enumerate(tb.rows):
        pr = row._tr.get_or_add_trPr()
        for el in list(pr):
            if el.tag in (qn('w:cantSplit'), qn('w:tblHeader')):
                pr.remove(el)
        pr.append(OxmlElement('w:cantSplit'))
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = i < len(tb.rows) - 1
                p.paragraph_format.keep_together = True
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                _no_indent(p)
                _no_grid(p)


def add_table(doc, headers, rows, description, number=None, widths=None, font_size=10):
    count = number if number is not None else _next_number(doc, r'^Table\s+(\d+)\s*$')
    captions = []
    for text in ('Table ' + str(count), description):
        p = add_body(doc, text)
        _no_indent(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        captions.append(p)
    tb = doc.add_table(rows=1, cols=len(headers))
    tb.autofit = False
    widths = [1] * len(headers) if widths is None else list(widths)
    if len(widths) != len(headers) or min(widths) <= 0:
        raise ValueError('Specify a positive relative width for each column')
    w = _available_width(doc)
    for col, weight in zip(tb.columns, widths):
        col.width = int(w * weight / sum(widths))
    for c, value in zip(tb.rows[0].cells, headers):
        c.text = str(value)
    for values in rows:
        if len(values) != len(headers):
            raise ValueError('Every row must match the header count')
        for c, value in zip(tb.add_row().cells, values):
            c.text = str(value)
    pr = tb._tbl.tblPr
    for child in list(pr):
        if child.tag in {qn('w:tblStyle'), qn('w:tblLook'), qn('w:tblBorders'), qn('w:shd')}:
            pr.remove(child)
    border = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single' if edge in ('top', 'bottom') else 'nil')
        if edge in ('top', 'bottom'):
            e.set(qn('w:sz'), '10')
            e.set(qn('w:color'), '000000')
        border.append(e)
    pr.append(border)
    for ri, row in enumerate(tb.rows):
        for ci, cell in enumerate(row.cells):
            cell.width = int(w * widths[ci] / sum(widths))
            cp = cell._tc.get_or_add_tcPr()
            for el in list(cp):
                if el.tag in (qn('w:shd'), qn('w:tcBorders')):
                    cp.remove(el)
            if ri == 0:
                cb = OxmlElement('w:tcBorders')
                line = OxmlElement('w:bottom')
                line.set(qn('w:val'), 'single')
                line.set(qn('w:sz'), '5')
                line.set(qn('w:color'), '000000')
                cb.append(line)
                cp.append(cb)
            for p in cell.paragraphs:
                p.style = style_for(doc, 'body')
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    _black(r)
                    r.font.size = Pt(font_size)
                    r.font.bold = False
                    r.font.italic = False
    keep_table_together(tb, captions)
    return tb


def add_figure(doc, image, description, number=None, width=None):
    count = number if number is not None else _next_number(doc, r'^Fig\.\s*(\d+)\.')
    p = doc.add_paragraph(style=style_for(doc, 'image'))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_indent(p)
    p.paragraph_format.keep_with_next = True
    width = min(_available_width(doc), width) if width is not None else _available_width(doc)
    p.add_run().add_picture(str(image), width=width)
    caption = add_body(doc, 'Fig.' + str(count) + '. ' + description)
    _no_indent(caption)
    caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
    caption.paragraph_format.keep_together = True
    return p, caption


def math_node(tag, children=(), **attrs):
    e = OxmlElement('m:' + tag)
    for k, v in attrs.items():
        e.set(qn('m:' + k), str(v))
    for child in children:
        e.append(child)
    return e


def math_run(text, upright=False):
    e = math_node('r')
    if upright:
        e.append(math_node('rPr', [math_node('sty', val='p')]))
    pr = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        fonts.set(qn('w:' + attr), 'Cambria Math')
    pr.append(fonts)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '000000')
    pr.append(color)
    e.append(pr)
    t = math_node('t')
    t.text = text
    e.append(t)
    return e


def fraction(numerator, denominator):
    return math_node('f', [math_node('num', numerator), math_node('den', denominator)])


def subscript(base, index):
    return math_node('sSub', [math_node('e', base), math_node('sub', index)])


def add_equation(doc, math, number=None):
    """Insert real OMML plus a right-tabbed SEQ field, not an image/LaTeX stub."""
    if isinstance(math, str):
        from docx.oxml import parse_xml
        math = parse_xml(math)
    omath = deepcopy(math) if hasattr(math, 'tag') else math_node('oMath', list(math))
    if omath.tag != qn('m:oMath'):
        raise ValueError('Supply an m:oMath element or its child elements')
    explicit_number = number is not None
    if number is None:
        previous = [f for f in doc.element.xpath('.//w:fldSimple')
                    if 'SEQ Equation' in f.get(qn('w:instr'), '')]
        cached = ''.join(t.text or '' for t in previous[-1].iter(qn('w:t'))) if previous else '0'
        number = int(cached) + 1 if cached.isdigit() else len(previous) + 1
    p = doc.add_paragraph(style=style_for(doc, 'body'))
    _no_indent(p)
    _no_grid(p)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.keep_together = True
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.tab_stops.add_tab_stop(_available_width(doc), WD_TAB_ALIGNMENT.RIGHT)
    p._p.append(omath)
    _black(p.add_run('\t('))
    f = OxmlElement('w:fldSimple')
    # Default field continues the document sequence; explicit numbering resets it.
    instruction = 'SEQ Equation \\* ARABIC'
    if explicit_number:
        instruction += ' \\r ' + str(number)
    f.set(qn('w:instr'), instruction)
    r, t = OxmlElement('w:r'), OxmlElement('w:t')
    t.text = str(number)
    r.append(t)
    f.append(r)
    p._p.append(f)
    _black(p.add_run(')'))
    return p


def check_document(path):
    """Structural checks, deliberately NOT a proof that tables fit on one page."""
    d = Document(str(path))
    issues = []
    for role in ALIASES:
        style_for(d, role)
    table_ids = []
    for i, tb in enumerate(d.tables, 1):
        caps = _table_captions(tb)
        if len(caps) != 2 or not all(p.paragraph_format.keep_with_next for p in caps):
            issues.append('Table ' + str(i) + ': caption pair is missing or detached')
        else:
            table_ids.append(int(caps[0].text.strip().split()[-1]))
        for ri, row in enumerate(tb.rows):
            if row._tr.get_or_add_trPr().find(qn('w:cantSplit')) is None:
                issues.append('Table ' + str(i) + ': row can split')
            if any(bool(p.paragraph_format.keep_with_next) != (ri < len(tb.rows) - 1)
                   for c in row.cells for p in c.paragraphs):
                issues.append('Table ' + str(i) + ': broken keep chain')
    if table_ids and table_ids != list(range(1, len(table_ids) + 1)):
        issues.append('Table numbers are not consecutive from 1')
    if not d.sections[0].different_first_page_header_footer:
        issues.append('First-page footer is not separate')
    if d.sections[0].first_page_footer._element.xpath('.//w:fldSimple'):
        issues.append('First-page footer contains a field')
    for p in d.paragraphs:
        if p._p.xpath('.//w:drawing'):
            if p.style.style_id != style_for(d, 'image').style_id or p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append('Image paragraph does not use centered TOC 1')
    result = {'tables': len(d.tables), 'images': len(d.inline_shapes),
              'native_equations': len(d.element.xpath('.//m:oMath')),
              'issues': list(dict.fromkeys(issues)), 'actual_pagination_checked': False}
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='command', required=True)
    new = sub.add_parser('new')
    new.add_argument('--output', type=Path, required=True)
    new.add_argument('--template', type=Path)
    check = sub.add_parser('check')
    check.add_argument('input', type=Path)
    args = p.parse_args()
    if args.command == 'new':
        if args.output.exists():
            raise FileExistsError('Refusing to overwrite an existing document')
        d = new_document(args.template)
        d.add_paragraph('', style=style_for(d, 'body'))
        d.save(args.output)
        print(str(args.output))
    else:
        report = check_document(args.input)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        if report['issues']:
            raise SystemExit(1)


if __name__ == '__main__':
    main()
