import io
import zipfile
from dataclasses import replace

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as Align
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from gostdoc.criteria import CATALOG, MISIS_PARAMETERS
from gostdoc.formatter import MISIS, format_docx, get_document
from gostdoc.preservation import assert_preserved
from gostdoc.workflow import audit_document, safe_docx_filename, zip_results


def save(doc):
    buffer = io.BytesIO(); doc.save(buffer); return buffer.getvalue()


def errors(data):
    return [i for i in audit_document(data, MISIS) if i['Уровень'] == 'Ошибка']


def test_catalog_is_bound_to_source_and_misis_profile():
    assert MISIS.summary() == MISIS_PARAMETERS
    assert CATALOG['source']['pages'] == 15
    assert len(CATALOG['source']['sha256']) == 64
    ids = [r['id'] for r in CATALOG['criteria']]
    assert len(ids) == len(set(ids))
    assert all(1 <= page <= 15 for rule in CATALOG['criteria'] for page in rule['pages'])
    with pytest.raises(ValueError):
        format_docx(save(Document()), replace(MISIS, size=18))


def test_body_second_run_and_inherited_spacing_are_audited():
    doc = Document(); p = doc.add_paragraph('Обычный текст. ')
    p.runs[0].font.name = 'Times New Roman'; p.runs[0].font.size = Pt(12)
    late = p.add_run('Второй фрагмент с ошибкой.'); late.font.name = 'Arial'; late.font.size = Pt(18)
    doc.styles['Normal'].paragraph_format.line_spacing = 2
    findings = errors(save(doc))
    assert any('Arial' in i['Замечание'] for i in findings)
    assert any('18' in i['Замечание'] for i in findings)
    assert any('Межстрочный' in i['Замечание'] for i in findings)
    fixed, *_ = format_docx(save(doc), MISIS)
    assert not errors(fixed)


def test_headings_context_bold_indentation_and_no_forced_page_breaks():
    doc = Document()
    doc.add_paragraph('Предыдущий текст.')
    first = doc.add_heading('Заголовок раздела', 1)
    second = doc.add_heading('Заголовок подраздела', 2)
    doc.add_paragraph('Основной текст.')
    fixed, *_ = format_docx(save(doc), MISIS)
    result = get_document(fixed); first, second = result.paragraphs[1:3]
    assert first.paragraph_format.space_before.pt == 24
    assert first.paragraph_format.space_after.pt == 0
    assert second.paragraph_format.space_after.pt == 24
    assert not first.paragraph_format.page_break_before
    assert not second.paragraph_format.page_break_before
    assert first.paragraph_format.first_line_indent.cm == pytest.approx(1.25, abs=.01)
    assert all(r.bold for r in first.runs if r.text)
    assert not errors(fixed)


def test_heading_at_explicit_page_start_and_broken_numbering_is_not_guessed():
    doc = Document(); doc.add_paragraph('Текст первой страницы.').add_run().add_break(__import__('docx').enum.text.WD_BREAK.PAGE)
    doc.add_heading('5 Непоследовательный номер', 1)
    doc.add_paragraph('Ссылка на раздел 5 остаётся верной.')
    fixed, before, after, warnings = format_docx(save(doc), MISIS)
    h = get_document(fixed).paragraphs[1]
    assert h.paragraph_format.space_before.pt == 0
    assert h.text == '5 Непоследовательный номер'
    assert before == after
    assert any('Нумерация' in warning for warning in warnings)


def test_consistent_heading_labels_become_real_multilevel_word_numbers():
    doc = Document()
    for text, level in [('1 Первый',1), ('1.1 Подраздел',2), ('1.2 Подраздел',2), ('2 Второй',1)]:
        doc.add_heading(text, level)
    doc.add_paragraph('См. раздел 1.2; эта фраза не переписывается.')
    fixed, *_ = format_docx(save(doc), MISIS)
    result = get_document(fixed)
    assert [p.text for p in result.paragraphs[:4]] == ['Первый','Подраздел','Подраздел','Второй']
    nums = [p._p.pPr.numPr for p in result.paragraphs[:4]]
    assert len({n.numId.val for n in nums}) == 1
    assert [n.ilvl.val for n in nums] == [0,1,1,0]
    assert result.paragraphs[4].text == 'См. раздел 1.2; эта фраза не переписывается.'
    assert '%1.%2' in result.part.numbering_part.element.xml


def test_forbidden_heading_is_reported_without_inventing_new_content():
    doc=Document();doc.add_heading('Ход работы',1)
    fixed,*_=format_docx(save(doc),MISIS)
    assert get_document(fixed).paragraphs[0].text == 'Ход работы'
    assert any('запрещает заголовок' in item['Замечание'] for item in errors(fixed))


def test_figure_caption_and_structural_label_changes_preserve_inline_objects():
    doc = Document(); heading = doc.add_paragraph('Оглавление')
    heading.add_run('')
    caption = doc.add_paragraph('Рисунок 1 '); caption.add_run('-'); caption.add_run(' Схема')
    bookmark = OxmlElement('w:bookmarkStart'); bookmark.set(qn('w:id'),'7'); bookmark.set(qn('w:name'),'Figure')
    caption._p.append(bookmark)
    fixed, *_ = format_docx(save(doc), MISIS)
    result = get_document(fixed)
    assert result.paragraphs[0].text == 'СОДЕРЖАНИЕ'
    assert result.paragraphs[0].alignment == Align.CENTER
    assert result.paragraphs[1].text == 'Рисунок 1 – Схема'
    assert result.paragraphs[1].alignment == Align.CENTER
    assert result.paragraphs[1]._p.xpath('.//w:bookmarkStart')
    assert result.sections[0].different_first_page_header_footer


def test_formula_layout_is_not_formatted_like_a_ten_point_table():
    doc = Document(); t = doc.add_table(rows=1, cols=3)
    math = OxmlElement('m:oMath'); mr = OxmlElement('m:r'); text = OxmlElement('m:t'); text.text = 'x+y'; mr.append(text); math.append(mr)
    t.cell(0,1).paragraphs[0]._p.append(math)
    t.cell(0,2).text = '(1)'
    fixed, *_ = format_docx(save(doc), MISIS)
    result = get_document(fixed); t = result.tables[0]
    assert t.cell(0,1).paragraphs[0].alignment == Align.CENTER
    assert t.cell(0,2).paragraphs[0].alignment == Align.RIGHT
    assert t.cell(0,2).paragraphs[0].runs[0].font.size.pt == 12
    assert t.cell(0,1)._tc.xpath('.//m:t')[0].text == 'x+y'
    assert not [i for i in errors(fixed) if i['Категория'] == 'Таблицы']


def test_existing_date_footer_does_not_prevent_page_field_creation():
    doc = Document(); doc.add_paragraph('Текст.')
    field = OxmlElement('w:fldSimple'); field.set(qn('w:instr'),' DATE ')
    doc.sections[0].footer.paragraphs[0]._p.append(field)
    fixed, *_ = format_docx(save(doc), MISIS)
    footer = get_document(fixed).sections[0].footer._element
    assert footer.xpath('.//w:fldSimple')[0].get(qn('w:instr')) == ' DATE '
    assert any('PAGE' in (n.text or '') for n in footer.xpath('.//w:instrText'))


def test_toc_styles_and_conditional_first_page_number_follow_pdf():
    doc = Document(); doc.add_paragraph('Введение.'); doc.add_paragraph('Основной текст.')
    fixed, *_ = format_docx(save(doc), MISIS, create_toc=True)
    result = get_document(fixed)
    title = next(p for p in result.paragraphs if p.text == 'СОДЕРЖАНИЕ')
    assert title.paragraph_format.space_after.pt == 24
    assert title.paragraph_format.line_spacing == 1.5
    assert result.sections[0].different_first_page_header_footer
    assert result.styles['TOC 2'].paragraph_format.left_indent.cm == pytest.approx(.62, abs=.01)
    assert result.styles['TOC 3'].paragraph_format.left_indent.cm == pytest.approx(1.25, abs=.01)
    assert result.styles['TOC 4'].paragraph_format.left_indent.cm == pytest.approx(1.875, abs=.01)
    intro = next(p for p in result.paragraphs if p.text == 'ВВЕДЕНИЕ')
    assert intro.style.name == 'Heading 1'
    assert any('"1-9"' in (n.text or '') for n in result.element.xpath('.//w:instrText'))


def test_existing_cover_is_frozen_separately_from_normal_body():
    doc = Document(); doc.styles['Normal'].font.name = 'Arial'; doc.styles['Normal'].font.size = Pt(14)
    doc.add_paragraph('Министерство науки · НИТУ МИСИС')
    doc.add_paragraph('Выполнил: студент, группа TEST-00')
    doc.add_paragraph('Москва').add_run().add_break(__import__('docx').enum.text.WD_BREAK.PAGE)
    doc.add_heading('ВВЕДЕНИЕ',1); doc.add_paragraph('Основной текст.')
    fixed, *_ = format_docx(save(doc), MISIS)
    result = get_document(fixed)
    assert result.paragraphs[0].style.font.name == 'Arial'
    assert result.paragraphs[0].style.font.size.pt == 14
    assert result.paragraphs[-1].runs[0].font.name == 'Times New Roman'
    assert result.paragraphs[-1].runs[0].font.size.pt == 12


@pytest.mark.parametrize('change', ['text','table','image','formula'])
def test_integrity_guard_refuses_content_loss(change):
    doc = Document(); doc.add_paragraph('Смысл работы: результат 42.')
    t = doc.add_table(rows=1, cols=1); t.cell(0,0).text = 'Данные'
    math = OxmlElement('m:oMath'); doc.paragraphs[0]._p.append(math)
    source = save(doc)
    if change == 'text':
        doc.paragraphs[0].runs[0].text = 'Изменённый смысл: результат 0.'
    elif change == 'table':
        t._tbl.getparent().remove(t._tbl)
    elif change == 'formula':
        math.getparent().remove(math)
    else:
        # A media package part cannot silently disappear, even when not exposed
        # by paragraph.text. Use a synthetic binary part in the package.
        src = io.BytesIO()
        with zipfile.ZipFile(src,'w') as output, zipfile.ZipFile(io.BytesIO(source)) as original:
            for name in original.namelist(): output.writestr(name,original.read(name))
            output.writestr('word/media/test.bin',b'original-image-bytes')
        source=src.getvalue()
    with pytest.raises(ValueError,match='заблокирован'):
        assert_preserved(source,save(doc))


def test_custom_filename_and_batch_names_are_safe_and_exact():
    assert safe_docx_filename('old.docx', custom_name='Моя.работа №1') == 'Моя.работа №1.docx'
    assert safe_docx_filename('old.docx', custom_name='Student_Group_PR№1.docx') == 'Student_Group_PR№1.docx'
    assert safe_docx_filename('old.docx', custom_name='../../CON.docx') == '_CON.docx'
    assert safe_docx_filename('old.docx', custom_name='') == 'old_GOST.docx'
    package=zip_results([('Мой отчёт.docx',b'a'),('Мой отчёт.docx',b'b')],custom_names=True)
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        assert archive.namelist() == ['Мой отчёт.docx','Мой отчёт_2.docx']
