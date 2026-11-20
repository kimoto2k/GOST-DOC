"""Paragraph roles shared by formatting and audit; source: misis_2024.json."""
import re
from dataclasses import dataclass

from docx.enum.text import WD_ALIGN_PARAGRAPH as Align
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .document import iter_runs, paragraph_value, run_value, style_chain

STRUCTURAL = {'ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ',
              'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', 'СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ'}


@dataclass(frozen=True)
class ParagraphRule:
    role: str
    page: int
    alignment: int = Align.JUSTIFY
    first: float = 1.25
    left: float = 0
    before: float = 0
    after: float = 0
    bold: bool = False


def heading_level(p, guess=False):
    name = p.style.name if p.style else ''
    match = re.fullmatch(r'(?:Heading|Заголовок)\s*([1-9])', name, re.I)
    if match:
        return int(match[1])
    if guess:
        from .formatter import is_heading
        return is_heading(p)
    return 0


def is_equation_table(table):
    return (len(table.rows) == 1 and len(table.columns) == 3
            and bool(table.cell(0, 1)._tc.xpath('.//m:oMath | .//m:oMathPara'))
            and bool(re.fullmatch(r'\s*\(\d+\)\s*', table.cell(0, 2).text)))


def is_page_start(p):
    if p.paragraph_format.page_break_before:
        return True
    previous = p._p.getprevious()
    return (previous is None or previous.tag == qn('w:p') and
            bool(previous.xpath('./w:pPr/w:sectPr | .//w:br[@w:type="page"]')))


def paragraph_rule(p, guess=False):
    from .formatter import is_list
    text = p.text.strip()
    normalized = text.rstrip('.').upper()
    name = p.style.name if p.style else ''
    toc = re.fullmatch(r'(?:TOC|Оглавление)\s*([1-9])', name, re.I)
    if toc:
        level = int(toc[1])
        left = {1: 0, 2: .62, 3: 1.25}.get(level, 1.25 + max(0, level - 3) * .625)
        return ParagraphRule('Содержание: запись', 13, Align.LEFT, first=0, left=left)
    if any(re.match(r'^\s*TOC(?:\s|$)', n.text or '', re.I) for n in p._p.xpath('.//w:instrText')):
        return ParagraphRule('Содержание: поле', 12, Align.LEFT, first=0)
    if normalized in STRUCTURAL or re.fullmatch(r'ПРИЛОЖЕНИЕ(?:\s+[А-ЯЁ])?', normalized):
        return ParagraphRule('Структурный заголовок', 12 if normalized in ('ОГЛАВЛЕНИЕ', 'СОДЕРЖАНИЕ') else 4,
                             Align.CENTER, first=0, before=0 if is_page_start(p) else 24, after=24, bold=True)
    level = heading_level(p, guess)
    if level:
        following = p._p.getnext()
        next_is_heading = False
        if following is not None and following.tag == qn('w:p'):
            from docx.text.paragraph import Paragraph
            next_is_heading = bool(heading_level(Paragraph(following, p._parent), guess))
        return ParagraphRule('Заголовок', 4, Align.LEFT, before=0 if is_page_start(p) else 24,
                             after=0 if next_is_heading else 24, bold=True)
    if p._p.xpath('.//m:oMath | .//m:oMathPara') and (not text or re.fullmatch(r'\(\d+\)', text)):
        return ParagraphRule('Формула', 10, Align.CENTER, first=0)
    if p._p.xpath('.//w:drawing | .//w:pict') and not text:
        return ParagraphRule('Рисунок', 5, Align.CENTER, first=0)
    if re.match(r'^Рисунок\s+(?:\d+|[А-ЯЁ]\.\d+)\s*[–—-]', text):
        return ParagraphRule('Подпись рисунка', 6, Align.CENTER, first=0)
    if is_list(p):
        return ParagraphRule('Список', 7, first=-1.25, left=1.25)
    return ParagraphRule('Основной текст', 5)


def format_paragraphs(doc, cfg, guess=False, cover_nodes=()):
    from .formatter import set_font
    from .cover import inside_cover
    for p in doc.paragraphs:
        if inside_cover(p._p, cover_nodes):
            continue
        if not p.text.strip() and not p._p.xpath('.//w:drawing | .//w:pict | .//m:oMath'):
            continue
        rule = paragraph_rule(p, guess)
        if rule.role == 'Заголовок' and not heading_level(p) and guess:
            p.style = f'Heading {heading_level(p, True)}'
        # Only literal formatting labels/punctuation from the PDF are changed.
        nodes = [n for n in p._p.xpath('.//w:t') if next(n.iterancestors(qn('w:p')), None) is p._p]
        visible = ''.join(n.text or '' for n in nodes)
        if visible.strip().rstrip('.').upper() == 'ОГЛАВЛЕНИЕ' and nodes:
            nodes[0].text = 'СОДЕРЖАНИЕ'
            for node in nodes[1:]:
                node.text = ''
        elif rule.role == 'Подпись рисунка':
            match = re.match(r'^Рисунок\s+\S+\s*([—-])', visible)
            if match:
                offset = match.start(1)
                for node in nodes:
                    value = node.text or ''
                    if offset < len(value):
                        node.text = value[:offset] + '–' + value[offset + 1:]
                        break
                    offset -= len(value)
        if rule.role == 'Структурный заголовок' and visible.rstrip().endswith('.'):
            for node in reversed(nodes):
                value = node.text or ''
                if value.strip():
                    node.text = value.rstrip().rstrip('.')
                    break
        if rule.role == 'Структурный заголовок':
            p.style = 'Normal' if p.text.strip().upper() == 'СОДЕРЖАНИЕ' else 'Heading 1'
            for node in nodes:
                node.text = (node.text or '').upper()
        pf = p.paragraph_format
        p.alignment = rule.alignment
        pf.left_indent = Cm(rule.left)
        pf.right_indent = Cm(0)
        pf.first_line_indent = Cm(rule.first)
        pf.space_before = Pt(rule.before)
        pf.space_after = Pt(rule.after)
        pf.line_spacing = cfg.line_spacing
        if rule.bold:
            pf.keep_with_next = True
        if p.text.strip().upper().startswith('ПРИЛОЖЕНИЕ') and rule.bold:
            pf.page_break_before = True
        for run in iter_runs(p):
            set_font(run, cfg, bold=True if rule.bold else None)
            if rule.bold:
                run.font.underline = False
            if rule.role == 'Структурный заголовок':
                # All caps displays uppercase without rewriting the student's text.
                run.font.all_caps = True


def format_equation_table(table, cfg):
    from .formatter import format_table_paragraph
    from docx.oxml import OxmlElement
    from docx.enum.table import WD_TABLE_ALIGNMENT
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, cell in enumerate(table.rows[0].cells):
        for p in cell.paragraphs:
            format_table_paragraph(p, cfg)
            p.alignment = (Align.LEFT, Align.CENTER, Align.RIGHT)[index]
        tcpr = cell._tc.get_or_add_tcPr()
        borders = tcpr.find(qn('w:tcBorders'))
        if borders is None:
            borders = OxmlElement('w:tcBorders'); tcpr.append(borders)
        for name in ('top','left','bottom','right','insideH','insideV'):
            border = borders.find(qn(f'w:{name}'))
            if border is None:
                border = OxmlElement(f'w:{name}'); borders.append(border)
            border.set(qn('w:val'), 'nil')


def _run_bool(run, p, attr):
    value = getattr(run.font, attr)
    if value is not None:
        return value
    for style in [*style_chain(run.style), *style_chain(p.style)]:
        value = getattr(style.font, attr)
        if value is not None:
            return value
    return False


def audit_paragraphs(doc, cfg):
    from .formatter import table_neighbors
    from .document import iter_tables
    from docx.shared import Length
    captions = {c._p for _, table in iter_tables(doc) for c, _ in [table_neighbors(table)] if c is not None}
    following = {p._p for _, table in iter_tables(doc) for _, p in [table_neighbors(table)] if p is not None}
    issues = []
    from .cover import cover_paragraphs
    protected = cover_paragraphs(doc)

    def add(pnum, rule, text, level='Ошибка'):
        issues.append({'Уровень': level, 'Категория': rule.role, 'Место': f'Абзац {pnum}',
                       'Замечание': f'{text} (методичка, стр. {rule.page}).'})

    for index, p in enumerate(doc.paragraphs, 1):
        if p._p in protected:
            continue
        if p._p in captions or not p.text.strip() and not p._p.xpath('.//w:drawing | .//m:oMath'):
            continue
        rule = paragraph_rule(p)
        if paragraph_value(p, 'alignment', doc) != rule.alignment:
            add(index, rule, f'Выравнивание: требуется {rule.alignment.name}')
        for attr, target, label in [('left_indent', rule.left, 'Отступ слева'),
                                    ('right_indent', 0, 'Отступ справа'),
                                    ('first_line_indent', rule.first, 'Первая строка')]:
            actual = paragraph_value(p, attr, doc).cm
            if abs(actual - target) > .03:
                add(index, rule, f'{label}: {actual:.2f} см; требуется {target:g} см')
        for attr, target, label in [('space_before', 6 if p._p in following else rule.before, 'перед'),
                                    ('space_after', rule.after, 'после')]:
            actual = paragraph_value(p, attr, doc).pt
            if abs(actual - target) > .1:
                add(index, rule, f'Интервал {label}: {actual:g} пт; требуется {target:g} пт')
        spacing = paragraph_value(p, 'line_spacing', doc)
        if isinstance(spacing, Length) or abs(spacing - cfg.line_spacing) > .03:
            add(index, rule, f'Межстрочный интервал должен быть {cfg.line_spacing:g}, не фиксированным в пт')
        names, sizes = set(), set()
        for run in iter_runs(p):
            if not run.text.strip():
                continue
            names.add(run_value(run, p, 'font', doc))
            sizes.add(run_value(run, p, 'size', doc))
            if rule.bold and not _run_bool(run, p, 'bold'):
                add(index, rule, 'Заголовок должен быть полужирным')
            if rule.bold and _run_bool(run, p, 'underline'):
                add(index, rule, 'Заголовок должен быть без подчёркивания')
            if rule.role == 'Структурный заголовок' and not run.text.isupper() and not _run_bool(run, p, 'all_caps'):
                add(index, rule, 'Структурный заголовок должен отображаться прописными буквами')
        if None in names or None in sizes:
            add(index, rule, 'Не определено эффективное значение шрифта/размера', 'Проверить вручную')
        if any(v is not None and v != cfg.font for v in names):
            add(index, rule, f'Требуется шрифт {cfg.font}; обнаружены {", ".join(sorted(str(v) for v in names))}')
        if any(v is not None and abs(v - cfg.size) > .2 for v in sizes):
            add(index, rule, f'Требуется размер {cfg.size:g} пт; обнаружены {", ".join(sorted(str(v) for v in sizes))}')
        if rule.role == 'Заголовок':
            if re.sub(r'^\d+(?:\.\d+)*\s+', '', p.text.strip()).upper() == 'ХОД РАБОТЫ':
                add(index, rule, 'Методичка запрещает заголовок «Ход работы»; задайте содержательное название')
            from .numbering import has_valid_heading_numbering
            if not has_valid_heading_numbering(doc, p, heading_level(p)):
                add(index, rule, 'Не подтверждена автоматическая многоуровневая нумерация с пробелом после номера')
            add(index, rule, 'Проверьте положение на настоящей странице и интервал до заголовка при автоматическом переносе', 'Проверить вручную')
        if rule.role == 'Основной текст' and p._p.xpath('.//m:oMath | .//m:oMathPara'):
            add(index, rule, 'Формула встроена в текст. По стр. 10 требуется отдельная строка; проверьте, не меняя смысл пояснения', 'Проверить вручную')
        if rule.role == 'Список':
            add(index, rule, 'Проверьте маркер «–», допустимые вложенные буквы/цифры и табуляцию списка', 'Проверить вручную')
        if p.text.strip().upper() == 'ОГЛАВЛЕНИЕ':
            add(index, rule, 'Название должно быть «СОДЕРЖАНИЕ»')
        if rule.role == 'Подпись рисунка' and re.match(r'^Рисунок\s+\S+\s*-', p.text.strip()):
            add(index, rule, 'В подписи требуется тире «–», а не дефис')
    # Reports must not multiply identical messages for differently styled runs.
    return [dict(zip(('Уровень','Категория','Место','Замечание'), values))
            for values in dict.fromkeys(tuple(i.values()) for i in issues)]
