"""Core DOCX formatter and limited rule checker for a GOST 7.32–2017 profile.

Does not claim full compliance or replace Word's pagination/layout engine.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import asdict, dataclass
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt, RGBColor

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

    def summary(self):
        return asdict(self)


STANDARD = Profile()


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
    title = insert_front_paragraph(doc, start_index, 'СОДЕРЖАНИЕ')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    title.paragraph_format.space_after = Pt(12)
    if begins_after_title:
        title.paragraph_format.page_break_before = True
    for run in title.runs:
        set_font(run, cfg, bold=True)
    content = insert_front_paragraph(doc, start_index + 1)
    content.paragraph_format.first_line_indent = Cm(0)
    r = content.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve')
    inst.text = ' TOC \\o "1-3" \\h \\z \\u '
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
        xml = footer._element.xml.upper()
        if ' PAGE ' in xml or 'W:FLDSIMPLE' in xml:
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

    for p in doc.paragraphs:
        level = is_heading(p) if recognize_headings else (
            1 if (p.style and (p.style.name or '').lower().startswith(('heading', 'заголов'))) else 0
        )
        process_paragraph(p, cfg, treat_as_heading=bool(level))

    if format_tables:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        process_paragraph(p, cfg, treat_as_heading=False, allow_style=False)
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
    add_page_numbers(doc, cfg, hide_on_first_page=add_title, warnings=warnings)

    after_text = extract_text(doc)
    output = io.BytesIO()
    doc.save(output)
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
        if not head and not is_list(p):
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
