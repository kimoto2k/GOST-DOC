"""UI-independent helpers: document summary, audit, profile import and batch ZIP.

The audit is intentionally partial: it inspects DOCX rules, not rendered pages.
"""
from __future__ import annotations

import csv
import io
import json
import re
import zipfile
from collections import Counter
from dataclasses import fields
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from .formatter import MISIS, Profile, STANDARD, diagnostics, get_document, is_heading, is_list, table_neighbors
from .document import iter_tables
from .table_audit import audit_tables

PROFILE_KEYS = {field.name for field in fields(Profile)}
MAX_BATCH_FILES = 10


def inspect_document(content: bytes) -> dict:
    """Return approximate text statistics; DOCX page count isn't available here."""
    doc = get_document(content)
    text = '\n'.join(p.text for p in doc.paragraphs if p.text.strip())
    table_words = sum(
        len(cell.text.split()) for table in doc.tables for row in table.rows for cell in row.cells
    )
    headings = [
        {'level': is_heading(p), 'text': p.text.strip()[:140]}
        for p in doc.paragraphs if p.text.strip() and is_heading(p)
    ]
    return {
        'paragraphs': sum(bool(p.text.strip()) for p in doc.paragraphs),
        'words': len(text.split()) + table_words,
        'tables': len(doc.tables),
        'sections': len(doc.sections),
        'headings': headings[:100],
        'headings_count': len(headings),
        'has_images': bool(doc.inline_shapes),
    }


def _inherited_attr(style, attr: str):
    """Resolve paragraph style property through its based-on chain, if possible."""
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        value = getattr(style.paragraph_format, attr, None)
        if value is not None:
            return value
        style = style.base_style
    return None


def _run_font(run, paragraph, doc):
    if run.font.name:
        return run.font.name
    style = paragraph.style
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        if style.font.name:
            return style.font.name
        style = style.base_style
    return doc.styles['Normal'].font.name


def _run_size(run, paragraph, doc):
    if run.font.size:
        return run.font.size.pt
    style = paragraph.style
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        if style.font.size:
            return style.font.size.pt
        style = style.base_style
    val = doc.styles['Normal'].font.size
    return val.pt if val else None


def audit_document(content: bytes, cfg: Profile = STANDARD) -> list[dict]:
    """Detailed but deliberately limited format audit, no full compliance claims.

    Fields are stable for CSV export and UI filtering.
    """
    doc = get_document(content)
    issues = []
    if cfg.ruleset == 'misis' and cfg != MISIS:
        raise ValueError('Параметры МИСИС должны совпадать с методичкой.')

    def add(level: str, kind: str, location: str, details: str):
        issues.append({
            'Уровень': level,
            'Категория': kind,
            'Место': location,
            'Замечание': details,
        })

    for item in diagnostics(content, cfg):
        detail = item['Проблема']
        if cfg.ruleset == 'misis' and not item['Элемент'].startswith('Раздел '):
            continue  # Role-aware checks below replace overlapping basic checks.
        if detail.startswith(('Левое', 'Правое', 'Верхнее', 'Нижнее', 'Формат')):
            add('Ошибка', 'Страница', item['Элемент'], detail)
        elif detail.startswith('Шрифт') or detail.startswith('Размер'):
            add('Ошибка', 'Типографика', item['Элемент'], detail)
        else:
            add('Замечание', 'Абзацы', item['Элемент'], detail)

    captions = {caption._p for _, table in iter_tables(doc)
                for caption, _ in [table_neighbors(table)] if caption is not None} if cfg.ruleset == 'misis' else set()
    for i, p in enumerate(doc.paragraphs, 1):
        if cfg.ruleset == 'misis':
            continue
        if not p.text.strip() or is_heading(p) or is_list(p) or p._p in captions:
            continue
        loc = f'Абзац {i}'
        sp = p.paragraph_format.line_spacing
        if sp is None and p.style:
            sp = _inherited_attr(p.style, 'line_spacing')
        if isinstance(sp, (int, float)) and abs(sp - cfg.line_spacing) > 0.03:
            add('Ошибка', 'Интервалы', loc, f'Межстрочный интервал {sp:g}, ожидается {cfg.line_spacing:g}.')
        indent = p.paragraph_format.first_line_indent
        if indent is None and p.style:
            indent = _inherited_attr(p.style, 'first_line_indent')
        if indent is not None and abs(indent.cm - cfg.first_line_cm) > 0.07:
            add('Замечание', 'Абзацы', loc, f'Красная строка {indent.cm:.2f} см вместо {cfg.first_line_cm:.2f} см.')
        for run in p.runs:
            if not run.text.strip():
                continue
            name = _run_font(run, p, doc)
            if name and name != cfg.font:
                add('Ошибка', 'Типографика', loc, f'Эффективный шрифт «{name}», требуется {cfg.font}.')
            size = _run_size(run, p, doc)
            if size is not None and abs(size - cfg.size) > 0.2:
                add('Ошибка', 'Типографика', loc, f'Эффективный размер {size:g} пт, ожидается {cfg.size}.')
            if run.font.color and run.font.color.rgb:
                if tuple(run.font.color.rgb) != (0, 0, 0):
                    add('Замечание', 'Цвет', loc, 'Текст не чёрный.')
            break

    if cfg.ruleset == 'misis':
        from .misis import audit_paragraphs
        issues.extend(audit_paragraphs(doc, cfg))
    issues.extend(audit_tables(doc, cfg))
    if doc.tables:
        add('Проверить вручную', 'Таблицы', 'Документ',
            f'Таблиц: {len(doc.tables)}. Проверь подписи, повтор заголовков и переносы между страницами.')
    if doc.inline_shapes:
        add('Проверить вручную', 'Иллюстрации', 'Документ',
            'Проверь номера, подписи, ссылки и положение рисунков.')
    if not any(is_heading(p) for p in doc.paragraphs):
        add('Проверить вручную', 'Структура', 'Документ',
            'Не распознаны заголовки. Применяй стили Word «Заголовок 1–3» для содержания.')
    if cfg.ruleset == 'misis':
        add('Проверить вручную', 'Методичка МИСИС', 'Документ',
            'Контекстные интервалы и нумерация заголовков, кафедральный титул, содержание, '
            'формулы, ссылки, источники и реальная вёрстка не проходят полный автоматический контроль.')
        if not any(re.search(r'\bцел(?:ь|ью|и)\b', p.text, re.I) for p in doc.paragraphs):
            add('Проверить вручную', 'Содержание работы', 'Документ',
                'Не найдена явно обозначенная цель работы. Проверьте требование стр. 5; цель не добавляется автоматически.')
        if doc.element.xpath('.//w:sdt'):
            add('Проверить вручную', 'Элементы управления Word', 'Документ',
                'Обнаружены элементы Rich Text/SDT. Их содержимое сохраняется; проверьте вложенные стили и титульник по стр. 14–15.')

    # Deduplicate when a basic diagnostic and a style-based check describe one issue.
    seen = set()
    unique = []
    for item in issues:
        key = (item['Категория'], item['Место'], item['Замечание'])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def audit_counts(items: list[dict]) -> dict[str, int]:
    counts = Counter(item['Уровень'] for item in items)
    return {
        'errors': counts['Ошибка'],
        'notes': counts['Замечание'],
        'manual': counts['Проверить вручную'],
    }


def audit_csv(items: list[dict]) -> bytes:
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=['Уровень', 'Категория', 'Место', 'Замечание'],
                            delimiter=';')
    writer.writeheader()
    writer.writerows(items)
    return ('\ufeff' + out.getvalue()).encode('utf-8')


def profile_to_json(profile: Profile) -> bytes:
    return json.dumps(profile.summary(), ensure_ascii=False, indent=2).encode('utf-8')


def profile_from_json(data: bytes) -> Profile:
    if len(data) > 16 * 1024:
        raise ValueError('Профиль слишком большой (максимум 16 КБ).')
    try:
        payload = json.loads(data.decode('utf-8-sig'))
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError('Некорректный JSON профиля.') from exc
    legacy_keys = PROFILE_KEYS - {'ruleset', 'table_size', 'table_line_spacing'}
    if not isinstance(payload, dict) or not legacy_keys <= set(payload) <= PROFILE_KEYS:
        raise ValueError('Файл должен содержать параметры профиля GOST DOC без неизвестных полей.')
    if payload.get('ruleset', 'generic') not in ('generic', 'misis', 'custom'):
        raise ValueError('Неизвестный набор требований профиля.')
    if not isinstance(payload['font'], str) or payload['font'] not in (
        'Times New Roman', 'Arial', 'Calibri',
    ):
        raise ValueError('Шрифт не поддерживается.')
    limits = {
        'size': (10, 18),
        'line_spacing': (1, 2),
        'first_line_cm': (0, 2.5),
        'left_cm': (0.5, 5),
        'right_cm': (0.5, 5),
        'top_cm': (0.5, 5),
        'bottom_cm': (0.5, 5),
    }
    for name, (low, high) in limits.items():
        value = payload[name]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not low <= value <= high:
            raise ValueError(f'Недопустимое значение параметра {name}.')
    for name, low, high in [('table_size', 8, 18), ('table_line_spacing', 1, 2)]:
        value = payload.get(name)
        if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float)) or not low <= value <= high):
            raise ValueError(f'Недопустимое значение параметра {name}.')
    profile = Profile(**payload)
    if profile.ruleset == 'misis' and profile != MISIS:
        raise ValueError('Профиль МИСИС фиксирован методичкой. Для других параметров выберите ruleset="custom".')
    return profile


def safe_docx_filename(name: str, *, custom_name: str | None = None) -> str:
    """Use a name without path components/control characters in ZIP exports."""
    custom = custom_name is not None and bool(custom_name.strip())
    name = custom_name if custom else name
    name = name.replace('\\', '/').split('/')[-1].strip()
    name = re.sub(r'[\x00-\x1f<>:"|?*]', '_', name)
    stem = ((name[:-5] if name.lower().endswith('.docx') else name) if custom else Path(name).stem)
    stem = stem.strip(' .')[:90] or 'document'
    if stem.upper() in ('CON', 'PRN', 'AUX', 'NUL') or re.fullmatch(r'(?:COM|LPT)[1-9]', stem.upper()):
        stem = '_' + stem
    return stem + ('' if custom else '_GOST') + '.docx'


def zip_results(files: list[tuple[str, bytes]], report: bytes | None = None, *, custom_names: bool = False) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        used = set()
        for name, data in files:
            filename = safe_docx_filename(name, custom_name=name if custom_names else None)
            stem = filename.removesuffix('.docx')
            n = 2
            while filename.casefold() in used:
                filename = f'{stem}_{n}.docx'
                n += 1
            used.add(filename.casefold())
            archive.writestr(filename, data)
        if report is not None:
            archive.writestr('audit_report.csv', report)
    return out.getvalue()


def demo_docx() -> bytes:
    """Small intentionally poorly formatted example, for safe testing."""
    doc = Document()
    doc.add_heading('ВВЕДЕНИЕ', level=1)
    p = doc.add_paragraph('В этом документе специально нарушено оформление: поля, интервалы и размер шрифта.')
    p.runs[0].font.size = __import__('docx').shared.Pt(16)
    doc.add_heading('1 Общие положения', level=1)
    doc.add_paragraph('Загрузите этот файл в GOST DOC и оцените найденные замечания.')
    doc.add_paragraph('Таблица 1 – Демонстрационные результаты')
    table = doc.add_table(rows=2, cols=2)
    for row, values in zip(table.rows, [('Показатель', 'Значение'), ('Тестовый результат', '42')]):
        for cell, value in zip(row.cells, values):
            cell.text = value
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.runs[0].font.size = __import__('docx').shared.Pt(12)
    doc.add_paragraph('Таким образом, проверка должна найти и исправить параметры таблицы.')
    section = doc.sections[0]
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()
