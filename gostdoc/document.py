"""Non-destructive DOCX traversal and effective formatting used by the audit."""
from docx.oxml.ns import qn
from docx.shared import Pt, Twips
from docx.text.run import Run
from lxml import etree


def iter_tables(parent, prefix=''):
    """Include nested tables, visiting each merged cell only once."""
    for index, table in enumerate(parent.tables, 1):
        location = f'{prefix}Таблица {index}'
        yield location, table
        for row, col, cell in iter_cells(table):
            yield from iter_tables(cell, f'{location}, строка {row}, ячейка {col} / ')


def iter_cells(table):
    seen = set()
    for row_index, row in enumerate(table.rows, 1):
        for col_index, cell in enumerate(row.cells, 1):
            if cell._tc not in seen:
                seen.add(cell._tc)
                yield row_index, col_index, cell


def iter_runs(paragraph):
    # paragraph.runs excludes hyperlinks. Never flatten paragraph.text: it removes
    # bookmarks, hyperlinks, fields, drawings and native equations.
    for element in paragraph._p.xpath('.//w:r'):
        if next(element.iterancestors(qn('w:p')), None) is paragraph._p:
            yield Run(element, paragraph)


def style_chain(style):
    seen = set()
    while style is not None and style.element not in seen:
        seen.add(style.element)
        yield style
        style = style.base_style


def paragraph_value(paragraph, attr, doc):
    value = getattr(paragraph.paragraph_format, attr)
    if value is not None:
        return value
    for style in style_chain(paragraph.style):
        value = getattr(style.paragraph_format, attr)
        if value is not None:
            return value
    defaults = doc.styles.element.xpath('./w:docDefaults/w:pPrDefault/w:pPr')
    ppr = defaults[0] if defaults else None
    if ppr is not None:
        if attr == 'line_spacing':
            nodes = ppr.xpath('./w:spacing')
            if nodes and nodes[0].get(qn('w:line')):
                line = int(nodes[0].get(qn('w:line')))
                return (line / 240 if nodes[0].get(qn('w:lineRule'), 'auto') == 'auto'
                        else Twips(line))
        mapping = {'space_before': ('spacing', 'before'), 'space_after': ('spacing', 'after'),
                   'first_line_indent': ('ind', 'firstLine'), 'left_indent': ('ind', 'left'),
                   'right_indent': ('ind', 'right')}
        if attr in mapping:
            node, name = mapping[attr]
            nodes = ppr.xpath(f'./w:{node}')
            if nodes and nodes[0].get(qn(f'w:{name}')):
                return Twips(int(nodes[0].get(qn(f'w:{name}'))))
    return 1.0 if attr == 'line_spacing' else Pt(0)


def _font_from_rpr(rpr, attr, doc):
    if rpr is None:
        return None
    if attr == 'size':
        nodes = rpr.xpath('./w:sz')
        return float(nodes[0].get(qn('w:val'))) / 2 if nodes else None
    nodes = rpr.xpath('./w:rFonts')
    if not nodes:
        return None
    fonts = nodes[0]
    theme = fonts.get(qn('w:asciiTheme')) or fonts.get(qn('w:hAnsiTheme'))
    if theme:
        for rel in doc.part.rels.values():
            if rel.reltype.endswith('/theme'):
                root = etree.fromstring(rel.target_part.blob)
                scheme = 'majorFont' if theme.startswith('major') else 'minorFont'
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                latin = root.find(f'a:themeElements/a:fontScheme/a:{scheme}/a:latin', ns)
                return latin.get('typeface') if latin is not None else None
    return fonts.get(qn('w:ascii')) or fonts.get(qn('w:hAnsi'))


def run_value(run, paragraph, attr, doc):
    for rpr in [run._r.rPr] + [s.element.rPr for s in style_chain(run.style)] + [
            s.element.rPr for s in style_chain(paragraph.style)]:
        value = _font_from_rpr(rpr, attr, doc)
        if value is not None:
            return value
    defaults = doc.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr')
    return _font_from_rpr(defaults[0] if defaults else None, attr, doc)
