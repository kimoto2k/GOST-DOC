"""Core DOCX formatter and limited rule checker for a GOST 7.32–2017 profile.

Does not claim full compliance or replace Word's pagination/layout engine.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import asdict, dataclass, replace
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt, RGBColor

from .document import iter_cells, iter_runs, iter_tables
from .criteria import MISIS_PARAMETERS

MAX_BYTES = 15 * 1024 * 1024
MAX_EXTRACTED_BYTES = 100 * 1024 * 1024
STRUCTURAL_TITLES = {
    'СПИСОК ИСПОЛНИТЕЛЕЙ', 'РЕФЕРАТ', 'СОДЕРЖАНИЕ',
    'ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ', 'ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ',
    'ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
    'ПРИЛОЖЕНИЕ',
}


@dataclass(frozen=True)
class Profile:
    font: str = 'Times New Roman'
    size: int = 12
    line_spacing: float = 1.5
    first_line_cm: float = 1.25
    left_cm: float = 3.0
    right_cm: float = 1.5
    top_cm: float = 2.0
    bottom_cm: float = 2.0
    ruleset: str = 'generic'
    table_size: Optional[float] = None
    table_line_spacing: Optional[float] = None

    def summary(self):
        return asdict(self)

    def for_table(self):
        return replace(self, size=self.table_size if self.table_size is not None else self.size,
                       line_spacing=self.table_line_spacing if self.table_line_spacing is not None else self.line_spacing)


STANDARD = Profile()
MISIS = Profile(**MISIS_PARAMETERS)


def validate_docx(content: bytes) -> None:
    if not content or len(content) > MAX_BYTES:
        raise ValueError('Файл пустой или превышает допустимые 15 МБ.')
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            parts = z.infolist()
            if len(parts) > 2500:
                raise ValueError('Слишком много компонентов внутри DOCX.')
            if sum(p.file_size for p in parts) > MAX_EXTRACTED_BYTES:
                raise ValueError('Слишком большой распакованный DOCX.')
            if 'word/document.xml' not in z.namelist():
                raise ValueError('В архиве нет Word-документа.')
            if any('vbaProject' in x for x in z.namelist()):
                raise ValueError('Документы с макросами не поддерживаются.')
            bad = z.testzip()
            if bad:
                raise ValueError('Повреждённый элемент DOCX.')
    except zipfile.BadZipFile as exc:
        raise ValueError('Это не корректный файл DOCX.') from exc


def get_document(content: bytes):
    validate_docx(content)
    return Document(io.BytesIO(content))


def set_font(style_or_run, profile: Profile, bold: Optional[bool] = None):
    font = style_or_run.font
    font.name = profile.font
    font.size = Pt(profile.size)
    font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        font.bold = bold
    if hasattr(style_or_run, '_element'):
        el = style_or_run._element
        rpr = el.get_or_add_rPr()
        rfonts = rpr.rFonts
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        for key in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            rfonts.set(qn(f'w:{key}'), profile.font)
        for key in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'csTheme', 'cstheme'):
            rfonts.attrib.pop(qn(f'w:{key}'), None)


def is_heading(paragraph) -> int:
    style_name = (paragraph.style.name or '').lower() if paragraph.style else ''
    m = re.search(r'(?:heading|заголовок)\s*([1-3])', style_name)
    if m:
        return int(m.group(1))
    text = paragraph.text.strip()
    if text.upper() in STRUCTURAL_TITLES:
        return 1
    m = re.match(r'^(\d+(?:\.\d+){0,2})\s+[^\s].*$', text)
    if m and len(text) < 125 and len(text.split()) < 15 and not text.endswith(('.', ':', ';')):
        return min(m.group(1).count('.') + 1, 3)
    return 0


def is_list(paragraph) -> bool:
    name = (paragraph.style.name or '').lower() if paragraph.style else ''
    if 'list' in name or 'список' in name:
        return True
    ppr = paragraph._p.pPr
    return bool(ppr is not None and ppr.numPr is not None)


def process_paragraph(paragraph, cfg: Profile, *, treat_as_heading: bool, allow_style: bool = True) -> None:
    if not paragraph.text.strip():
        return
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = cfg.line_spacing

    if treat_as_heading:
        level = is_heading(paragraph)
        text = paragraph.text.strip()
        structural = text.upper() in STRUCTURAL_TITLES
        if allow_style and level:
            paragraph.style = f'Heading {level}'
        pf.keep_with_next = True
        pf.widow_control = True
        if structural:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.page_break_before = True
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(cfg.first_line_cm)
            if level == 1:
                pf.page_break_before = True
        for run in paragraph.runs:
            set_font(run, cfg, bold=True)
            run.font.underline = False
    elif is_list(paragraph):
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            set_font(run, cfg)
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(cfg.first_line_cm)
        for run in paragraph.runs:
            set_font(run, cfg)


def format_table_paragraph(paragraph, cfg):
    """Table-specific direct properties also protect empty cells from Normal."""
    pf = paragraph.paragraph_format
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.left_indent = pf.right_indent = pf.first_line_indent = Cm(0)
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = cfg.line_spacing
    for run in iter_runs(paragraph):
        set_font(run, cfg)
    # Empty cells need a paragraph-mark size too, for text entered later in Word.
    ppr = paragraph._p.get_or_add_pPr()
    rpr = ppr.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        ppr.insert_element_before(rpr, 'w:sectPr', 'w:pPrChange')
    for name, value in [('rFonts', cfg.font), ('sz', str(int(cfg.size * 2)))]:
        node = rpr.find(qn(f'w:{name}'))
        if node is None:
            node = OxmlElement(f'w:{name}')
            rpr.append(node)
        if name == 'rFonts':
            for key in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                node.set(qn(f'w:{key}'), value)
            for key in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'csTheme', 'cstheme'):
                node.attrib.pop(qn(f'w:{key}'), None)
        else:
            node.set(qn('w:val'), value)


def table_neighbors(table):
    """Only unambiguous adjacent captions/body text; don't invent missing text."""
    # A caption inside a parent cell is still table content (10 pt), not body
    # text (12 pt). Do not impose conflicting surrounding-body rules on it.
    if table._tbl.getparent().tag != qn('w:body'):
        return None, None
    previous, following = table._tbl.getprevious(), table._tbl.getnext()
    caption = None
    if previous is not None and previous.tag == qn('w:p'):
        p = Paragraph(previous, table._parent)
        if re.match(r'^(?:Таблица|Продолжение таблицы)\s+(?:\d+|[А-ЯЁ]\.\d+)(?:\s|[–—.-]|$)', p.text.strip()):
            caption = p
    after = None
    if following is not None and following.tag == qn('w:p'):
        p = Paragraph(following, table._parent)
        if p.text.strip() and not is_heading(p) and not is_list(p) and not re.match(
                r'^(?:Рисунок|Таблица|Продолжение таблицы)\s', p.text.strip()):
            after = p
    return caption, after


def format_table(table, cfg):
    from .misis import is_equation_table, format_equation_table
    if cfg.ruleset == 'misis' and is_equation_table(table):
        format_equation_table(table, cfg)
        return
    table_cfg = cfg.for_table()
    if cfg.ruleset == 'misis':
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        width = table._tbl.tblPr.find(qn('w:tblW'))
        if width is None:
            width = OxmlElement('w:tblW')
            table._tbl.tblPr.append(width)
        width.set(qn('w:type'), 'pct')
        width.set(qn('w:w'), '5000')
    for _, _, cell in iter_cells(table):
        if cfg.ruleset == 'misis':
            tcpr = cell._tc.get_or_add_tcPr()
            shading = tcpr.find(qn('w:shd'))
            if shading is None:
                shading = OxmlElement('w:shd')
                tcpr.append(shading)
            shading.set(qn('w:val'), 'clear')
            shading.set(qn('w:fill'), 'auto')
            shading.set(qn('w:color'), 'auto')
            for key in ('themeFill', 'themeFillTint', 'themeFillShade', 'themeColor', 'themeTint', 'themeShade'):
                shading.attrib.pop(qn(f'w:{key}'), None)
        for p in cell.paragraphs:
            format_table_paragraph(p, table_cfg)
    if cfg.ruleset == 'misis':
        caption, after = table_neighbors(table)
        if caption is not None:
            format_table_paragraph(caption, cfg)
        if after is not None:
            after.paragraph_format.space_before = Pt(6)


def ensure_page_field(paragraph) -> None:
    r = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    r._r.extend([begin, instr, end])


def update_on_open(doc) -> None:
    settings = doc.settings.element
    u = settings.find(qn('w:updateFields'))
    if u is None:
        u = OxmlElement('w:updateFields')
        settings.append(u)
    u.set(qn('w:val'), 'true')


def insert_front_paragraph(doc, index: int, text: str = ''):
    paragraph = doc.add_paragraph()
    if text:
        paragraph.add_run(text)
    doc.element.body.insert(index, paragraph._p)
    return paragraph


def title_page(doc, cfg: Profile, data: dict) -> int:
    # Simplified student cover, NOT the formal Appendix A NIR report.
    elements = [
        ('org', data.get('organization', ''), 'center', False),
        ('university', data.get('university', ''), 'center', True),
        ('department', data.get('department', ''), 'center', False),
        ('space', '', 'center', False),
        ('space', '', 'center', False),
        ('type', data.get('work_type', 'ОТЧЁТ'), 'center', True),
        ('topic', data.get('topic', 'Тема работы'), 'center', True),
        ('space', '', 'center', False),
        ('student', f"Выполнил(а): {data.get('student', '')}", 'right', False),
        ('group', f"Группа: {data.get('group', '')}", 'right', False),
        ('teacher', f"Проверил(а): {data.get('teacher', '')}", 'right', False),
        ('space', '', 'center', False),
        ('city', f"{data.get('city', 'Москва')} {data.get('year', '')}", 'center', False),
    ]
    for idx, (kind, value, alignment, bold) in enumerate(elements):
        p = insert_front_paragraph(doc, idx, value)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_after = Pt(8 if kind != 'space' else 16)
        p.paragraph_format.line_spacing = 1.0
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if alignment == 'center' else WD_ALIGN_PARAGRAPH.RIGHT
        for run in p.runs:
            set_font(run, cfg, bold=bold)
    return len(elements)


def add_toc(doc, cfg: Profile, start_index: int, *, begins_after_title: bool = False) -> None:
    if cfg.ruleset == 'misis':
        from docx.enum.style import WD_STYLE_TYPE
        for level in range(1, 10):
            style = next((s for s in doc.styles if re.fullmatch(r'(?:TOC|Оглавление)\s*' + str(level), s.name, re.I)), None)
            if style is None:
                style = doc.styles.add_style(f'TOC {level}', WD_STYLE_TYPE.PARAGRAPH)
            set_font(style, cfg)
            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm({1: 0, 2: .62, 3: 1.25}.get(level, 1.25 + max(0, level - 3) * .625))
            pf.right_indent = pf.first_line_indent = Cm(0)
            pf.space_before = pf.space_after = Pt(0)
            pf.line_spacing = cfg.line_spacing
    title = insert_front_paragraph(doc, start_index, 'СОДЕРЖАНИЕ')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(24 if cfg.ruleset == 'misis' else 12)
    title.paragraph_format.line_spacing = cfg.line_spacing
    if begins_after_title:
        title.paragraph_format.page_break_before = True
    for run in title.runs:
        set_font(run, cfg, bold=True)
    content = insert_front_paragraph(doc, start_index + 1)
    content.paragraph_format.first_line_indent = Cm(0)
    content.alignment = WD_ALIGN_PARAGRAPH.LEFT
    content.paragraph_format.space_before = content.paragraph_format.space_after = Pt(0)
    content.paragraph_format.line_spacing = cfg.line_spacing
    r = content.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve')
    levels = '1-9' if cfg.ruleset == 'misis' else '1-3'
    inst.text = f' TOC \\o "{levels}" \\h \\z \\u '
    separate = OxmlElement('w:fldChar'); separate.set(qn('w:fldCharType'), 'separate')
    placeholder = OxmlElement('w:t'); placeholder.text = 'В Word: щёлкните правой кнопкой и выберите «Обновить поле».'
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    r._r.extend([begin, inst, separate, placeholder, end])
    set_font(r, cfg)
    update_on_open(doc)


def add_page_numbers(doc, cfg: Profile, hide_on_first_page: bool, warnings: list[str]) -> None:
    for index, section in enumerate(doc.sections):
        if hide_on_first_page and index == 0:
            section.different_first_page_header_footer = True
            if any(p.text.strip() or p._p.xpath('.//w:fldChar') for p in section.first_page_footer.paragraphs):
                warnings.append('На титуле был существующий колонтитул: проверьте отсутствие номера вручную.')
        footer = section.footer
        instructions = [n.text or '' for n in footer._element.xpath('.//w:instrText')]
        instructions += [n.get(qn('w:instr'), '') for n in footer._element.xpath('.//w:fldSimple')]
        if any(re.match(r'^\s*PAGE(?:\s|$)', value, re.I) for value in instructions):
            warnings.append(f'Раздел {index+1}: уже есть поле в нижнем колонтитуле, нумерация сохранена.')
            continue
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        ensure_page_field(p)
        for run in p.runs:
            set_font(run, cfg)


def format_docx(content: bytes, cfg: Profile = STANDARD, *, add_title: bool = False,
                cover: Optional[dict] = None, create_toc: bool = False,
                recognize_headings: bool = True, format_tables: bool = True):
    doc = get_document(content)
    original_first = doc.element.body[0] if len(doc.element.body) else None
    before_text = extract_text(doc)
    warnings: list[str] = []
    cover_nodes = set()
    numbered_positions = set()
    if cfg.ruleset == 'misis':
        if cfg != MISIS:
            raise ValueError('Параметры МИСИС должны совпадать с методичкой. Для других значений используйте свой профиль.')
        from .cover import protect_cover
        cover_nodes = protect_cover(doc)
        if cover_nodes:
            warnings.append('Найден исходный титульный лист: его стили защищены от изменения основного текста. Сравните с загруженным эталоном.')
        warnings.append('МИСИС: контекстные интервалы и нумерацию заголовков, кафедральный титул, '
                        'содержание, ссылки, формулы и вёрстку проверьте в Word. Это не полный нормоконтроль.')

    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(cfg.left_cm)
        section.right_margin = Cm(cfg.right_cm)
        section.top_margin = Cm(cfg.top_cm)
        section.bottom_margin = Cm(cfg.bottom_cm)

    normal = doc.styles['Normal']
    set_font(normal, cfg)
    normal.paragraph_format.line_spacing = cfg.line_spacing
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.first_line_indent = Cm(cfg.first_line_cm)
    for name in ('Heading 1', 'Heading 2', 'Heading 3'):
        style = doc.styles[name]
        set_font(style, cfg, bold=True)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_after = Pt(0)

    if cfg.ruleset == 'misis':
        from .misis import format_paragraphs
        format_paragraphs(doc, cfg, guess=recognize_headings, cover_nodes=cover_nodes)
        from .numbering import auto_number_headings
        if not auto_number_headings(doc, guess=recognize_headings, excluded=cover_nodes, changed_positions=numbered_positions):
            warnings.append('Нумерация заголовков не перенесена в автоматическую: обнаружены пропуски или несовпадение уровней. Проверьте, чтобы не исказить ссылки на разделы.')
    else:
        for p in doc.paragraphs:
            level = is_heading(p) if recognize_headings else (
                1 if (p.style and (p.style.name or '').lower().startswith(('heading', 'заголов'))) else 0
            )
            process_paragraph(p, cfg, treat_as_heading=bool(level))

    if format_tables:
        for _, table in iter_tables(doc):
            from .cover import inside_cover
            if inside_cover(table._tbl, cover_nodes):
                continue
            format_table(table, cfg)
    if doc.tables:
        warnings.append('Проверьте таблицы: ширину, переносы, подписи, объединённые ячейки и разбиение по страницам.')
    if doc.inline_shapes:
        warnings.append('Проверьте расположение изображений и подписей к рисункам.')

    start = title_page(doc, cfg, cover or {}) if add_title else 0
    if create_toc:
        add_toc(doc, cfg, start, begins_after_title=add_title)
        warnings.append('Для содержания в Microsoft Word требуется обновление поля (F9 / «Обновить поле»).')
    if (add_title or create_toc) and original_first is not None:
        if original_first.tag == qn('w:p'):
            Paragraph(original_first, doc._body).paragraph_format.page_break_before = True
        else:
            boundary = insert_front_paragraph(doc, start + (2 if create_toc else 0))
            boundary.paragraph_format.page_break_before = True
    has_toc = create_toc or any(p.text.strip().upper() in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ') for p in doc.paragraphs)
    has_toc = has_toc or any(re.match(r'^\s*TOC(?:\s|$)', n.text or '', re.I)
                             for n in doc.element.xpath('.//w:instrText'))
    add_page_numbers(doc, cfg, hide_on_first_page=has_toc if cfg.ruleset == 'misis' else add_title, warnings=warnings)

    after_text = extract_text(doc)
    output = io.BytesIO()
    doc.save(output)
    from .preservation import assert_preserved
    assert_preserved(content, output.getvalue(), heading_positions=numbered_positions)
    return output.getvalue(), before_text, after_text, warnings


def extract_text(doc) -> str:
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            parts.append(' | '.join(c.text.replace('\n', ' ') for c in row.cells))
    return '\n'.join(parts)


def diagnostics(content: bytes, cfg: Profile = STANDARD) -> list[dict]:
    doc = get_document(content)
    found = []
    captions = {caption._p for _, table in iter_tables(doc)
                for caption, _ in [table_neighbors(table)] if caption is not None} if cfg.ruleset == 'misis' else set()
    for i, sec in enumerate(doc.sections, 1):
        for key, actual, expected in (
            ('Левое поле', sec.left_margin.cm, cfg.left_cm),
            ('Правое поле', sec.right_margin.cm, cfg.right_cm),
            ('Верхнее поле', sec.top_margin.cm, cfg.top_cm),
            ('Нижнее поле', sec.bottom_margin.cm, cfg.bottom_cm),
        ):
            if abs(actual-expected) > 0.06:
                found.append({'Элемент': f'Раздел {i}', 'Проблема': f'{key}: {actual:.2f} см вместо {expected:.2f} см'})
        if abs(sec.page_width.cm-21)>0.06 or abs(sec.page_height.cm-29.7)>0.06:
            found.append({'Элемент': f'Раздел {i}', 'Проблема': 'Формат страницы отличается от A4'})
    for i, p in enumerate(doc.paragraphs, 1):
        if not p.text.strip():
            continue
        head = bool(is_heading(p))
        if not head and not is_list(p) and p._p not in captions:
            if p.alignment not in (None, WD_ALIGN_PARAGRAPH.JUSTIFY):
                found.append({'Элемент': f'Абзац {i}', 'Проблема': 'Выравнивание не по ширине'})
            if p.paragraph_format.first_line_indent is not None:
                if abs(p.paragraph_format.first_line_indent.cm-cfg.first_line_cm) > 0.07:
                    found.append({'Элемент': f'Абзац {i}', 'Проблема': 'Красная строка отличается от профиля'})
        for run in p.runs:
            if not run.text.strip():
                continue
            if run.font.name and run.font.name != cfg.font:
                found.append({'Элемент': f'Абзац {i}', 'Проблема': f'Шрифт: {run.font.name}'})
                break
            if run.font.size and abs(run.font.size.pt-cfg.size) > 0.2:
                found.append({'Элемент': f'Абзац {i}', 'Проблема': f'Размер: {run.font.size.pt:g} пт'})
                break
    return found
