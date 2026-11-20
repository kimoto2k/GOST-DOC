import io
import pytest
from docx import Document
from docx.shared import Pt

from gostdoc.formatter import (
    Profile, STANDARD, diagnostics, format_docx, get_document, validate_docx,
)


def sample_doc():
    doc = Document()
    doc.add_paragraph('ВВЕДЕНИЕ')
    p = doc.add_paragraph('Неправильно оформленный основной текст.')
    p.add_run(' Продолжение').font.size = Pt(16)
    doc.add_paragraph('1 Теоретические основы')
    doc.add_paragraph('Текст исследования с обычным содержимым.')
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = 'Показатель'
    t.cell(0, 1).text = 'Значение'
    b = io.BytesIO()
    doc.save(b)
    return b.getvalue()


def test_invalid():
    with pytest.raises(ValueError):
        validate_docx(b'wrong')


def test_base_formatting_preserves_text_and_tables():
    data = sample_doc()
    output, before, after, warnings = format_docx(data)
    doc = get_document(output)
    assert doc.sections[0].left_margin.cm == pytest.approx(3.0, abs=.06)
    assert doc.sections[0].right_margin.cm == pytest.approx(1.5, abs=.06)
    assert doc.sections[0].page_width.cm == pytest.approx(21, abs=.06)
    assert len(doc.tables) == 1
    assert doc.tables[0].cell(0, 0).text == 'Показатель'
    assert before == after
    assert 'Текст исследования' in after
    assert any('таблицы' in x.lower() for x in warnings)


def test_title_toc_page_and_field_structure():
    output, before, after, warnings = format_docx(
        sample_doc(), add_title=True, cover={
            'organization': 'Министерство', 'university': 'Университет',
            'department': 'Кафедра', 'work_type': 'ОТЧЁТ', 'topic': 'Тема',
            'student': 'Иван Иванов', 'group': '01', 'teacher': 'Петров',
            'city': 'Москва', 'year': '2026',
        }, create_toc=True)
    doc = get_document(output)
    texts = [p.text for p in doc.paragraphs]
    assert texts.index('Университет') < texts.index('СОДЕРЖАНИЕ') < texts.index('ВВЕДЕНИЕ')
    assert doc.sections[0].different_first_page_header_footer
    xml = '\n'.join(p._p.xml for p in doc.paragraphs)
    assert 'TOC' in xml
    assert 'PAGE' in doc.sections[0].footer._element.xml
    assert any('обновлен' in w.lower() or 'обнов' in w.lower() for w in warnings)


def test_diagnostics_and_custom_profile():
    data = sample_doc()
    original = diagnostics(data)
    assert original
    custom = Profile(size=13, right_cm=2)
    output, *_ = format_docx(data, custom)
    doc = get_document(output)
    assert doc.sections[0].right_margin.cm == pytest.approx(2, abs=.06)
    assert doc.paragraphs[3].runs[0].font.size.pt == pytest.approx(13)
