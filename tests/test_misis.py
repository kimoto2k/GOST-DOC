"""Regression documents contain synthetic content only; inspect saved DOCX XML."""
import io
import json
import zipfile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image
from lxml import etree

from gostdoc.document import iter_cells, iter_runs, iter_tables
from gostdoc.formatter import MISIS, STANDARD, Profile, format_docx, get_document
from gostdoc.workflow import audit_document, profile_from_json, profile_to_json


def save(doc):
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def sample():
    doc = Document()
    doc.add_heading('ВВЕДЕНИЕ', level=1)
    doc.add_paragraph('Цель работы — проверить оформление синтетического документа.')
    doc.add_paragraph('Таблица 1 – Результаты')
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = 'Показатель'
    table.cell(0, 1).text = 'Значение'
    table.cell(1, 0).text = 'Данные'
    table.cell(1, 1).text = '42'
    table.cell(2, 0).merge(table.cell(2, 1)).text = 'Объединённая ячейка'
    for _, _, cell in iter_cells(table):
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf = p.paragraph_format
            pf.first_line_indent = pf.left_indent = pf.right_indent = Cm(1.25)
            pf.space_before = pf.space_after = Pt(8)
            pf.line_spacing = 1.5
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(12)
    doc.add_paragraph('Таким образом, результаты подтверждают вывод.')
    return doc


def table_errors(data, cfg=MISIS):
    return [issue for issue in audit_document(data, cfg)
            if issue['Категория'] == 'Таблицы' and issue['Уровень'] == 'Ошибка']


def test_misis_tables_round_trip_and_body_separation():
    source = save(sample())
    output, before, after, _ = format_docx(source, MISIS)
    result = get_document(output)
    assert before == after
    assert result.paragraphs[1].runs[0].font.size.pt == 12
    assert result.paragraphs[1].paragraph_format.line_spacing == 1.5
    table = result.tables[0]
    assert table.alignment == WD_TABLE_ALIGNMENT.CENTER
    width = table._tbl.tblPr.find(qn('w:tblW'))
    assert (width.get(qn('w:type')), width.get(qn('w:w'))) == ('pct', '5000')
    assert len(table.rows) == 3
    assert table.cell(2, 0)._tc is table.cell(2, 1)._tc
    for _, _, cell in iter_cells(table):
        for p in cell.paragraphs:
            assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
            assert p.paragraph_format.line_spacing == 1
            assert p.paragraph_format.first_line_indent.cm == 0
            assert p.paragraph_format.left_indent.cm == 0
            assert p.paragraph_format.right_indent.cm == 0
            assert p.paragraph_format.space_before.pt == 0
            assert p.paragraph_format.space_after.pt == 0
            for run in iter_runs(p):
                assert run.font.name == 'Times New Roman'
                assert run.font.size.pt == 10
    assert table_errors(source)
    assert not table_errors(output)


def test_caption_and_following_text_have_distinct_parameters():
    output, *_ = format_docx(save(sample()), MISIS)
    paragraphs = get_document(output).paragraphs
    caption, after = paragraphs[2], paragraphs[3]
    assert caption.text == 'Таблица 1 – Результаты'
    assert caption.runs[0].font.size.pt == 12
    assert caption.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert caption.paragraph_format.line_spacing == 1.5
    assert caption.paragraph_format.first_line_indent.cm == 0
    assert after.paragraph_format.space_before.pt == 6
    assert after.paragraph_format.first_line_indent.cm == pytest.approx(1.25, abs=.01)
    assert not any('Красная строка' in item['Замечание'] and item['Место'] == 'Абзац 3'
                   for item in audit_document(output, MISIS))


def test_nested_and_empty_cells_are_formatted_once():
    doc = sample()
    cell = doc.tables[0].cell(1, 0)
    cell.add_paragraph('Таблица 1.1 – Вложенные данные')
    inner = cell.add_table(rows=1, cols=2)
    inner.cell(0, 0).text = 'Вложенный текст'
    cell.paragraphs[-1].text = 'Вывод внутри внешней ячейки.'
    output, *_ = format_docx(save(doc), MISIS)
    tables = list(iter_tables(get_document(output)))
    assert len(tables) == 2
    empty = tables[1][1].cell(0, 1).paragraphs[0]
    assert not empty.text
    assert empty.paragraph_format.first_line_indent.cm == 0
    assert empty.paragraph_format.line_spacing == 1
    assert empty._p.xpath('./w:pPr/w:rPr/w:sz')[0].get(qn('w:val')) == '20'
    assert tables[0][1].cell(1, 0).paragraphs[-1].paragraph_format.space_before.pt == 0
    assert not table_errors(output)


def test_audit_checks_every_run_character_style_hyperlink_and_exact_spacing():
    doc = sample()
    p = doc.tables[0].cell(0, 0).paragraphs[0]
    p.runs[0].font.name = 'Times New Roman'
    p.runs[0].font.size = Pt(10)
    style = doc.styles.add_style('Wrong cell character', WD_STYLE_TYPE.CHARACTER)
    style.font.size = Pt(17)
    p.add_run(' Второй фрагмент').style = style
    hyperlink = OxmlElement('w:hyperlink')
    r = OxmlElement('w:r')
    rp = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts'); fonts.set(qn('w:ascii'), 'Calibri')
    rp.append(fonts); r.append(rp)
    text = OxmlElement('w:t'); text.text = 'Ссылка'
    r.append(text); hyperlink.append(r); p._p.append(hyperlink)
    p.paragraph_format.line_spacing = Pt(12)
    issues = table_errors(save(doc))
    assert any('17 пт' in i['Замечание'] for i in issues)
    assert any('Calibri' in i['Замечание'] for i in issues)
    assert any('фиксированный' in i['Замечание'] for i in issues)
    fixed, *_ = format_docx(save(doc), MISIS)
    assert not table_errors(fixed)
    assert 'Ссылка' in get_document(fixed).tables[0].cell(0, 0).text


def test_inherited_bad_paragraph_style_is_detected_and_overridden():
    doc = sample()
    style = doc.styles.add_style('Inherited table paragraph', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'; style.font.size = Pt(15)
    style.paragraph_format.line_spacing = 2
    style.paragraph_format.first_line_indent = Cm(.5)
    p = doc.tables[0].cell(1, 0).paragraphs[0]
    p.style = style
    p.paragraph_format.line_spacing = None
    p.paragraph_format.first_line_indent = None
    p.runs[0].font.name = None; p.runs[0].font.size = None
    issues = table_errors(save(doc))
    assert any('15 пт' in i['Замечание'] for i in issues)
    assert any('Межстрочный интервал 2' in i['Замечание'] for i in issues)
    assert any('Красная строка: 0.50' in i['Замечание'] for i in issues)
    fixed, *_ = format_docx(save(doc), MISIS)
    assert not table_errors(fixed)


def test_theme_font_overrides_are_removed_for_misis_table_runs():
    doc = sample()
    theme = next(rel.target_part for rel in doc.part.rels.values() if rel.reltype.endswith('/theme'))
    root = etree.fromstring(theme.blob)
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    root.find('a:themeElements/a:fontScheme/a:minorFont/a:latin', ns).set('typeface', 'Arial')
    theme._blob = etree.tostring(root)
    run = doc.tables[0].cell(0, 0).paragraphs[0].runs[0]
    fonts = run._r.rPr.rFonts
    fonts.set(qn('w:ascii'), 'Times New Roman')
    fonts.set(qn('w:asciiTheme'), 'minorHAnsi')
    assert any('Шрифт: Arial' in i['Замечание'] and 'строка 1, ячейка 1' in i['Место']
               for i in table_errors(save(doc)))
    fixed, *_ = format_docx(save(doc), MISIS)
    fonts = get_document(fixed).tables[0].cell(0, 0).paragraphs[0].runs[0]._r.rPr.rFonts
    assert fonts.get(qn('w:ascii')) == 'Times New Roman'
    assert fonts.get(qn('w:asciiTheme')) is None
    assert not table_errors(fixed)


def test_shading_is_cleared_without_removing_number_column_or_header_text():
    doc = sample(); cell = doc.tables[0].cell(0, 0); cell.text = '№ п/п'
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'FF0000'); shd.set(qn('w:themeFill'), 'accent1')
    cell._tc.get_or_add_tcPr().append(shd)
    source = save(doc)
    assert any('заливка' in i['Замечание'] for i in table_errors(source))
    output, *_ = format_docx(source, MISIS)
    result = get_document(output)
    assert result.tables[0].cell(0, 0).text == '№ п/п'
    assert len(result.tables[0].columns) == 2
    assert not table_errors(output)
    assert any('порядковым номером' in i['Замечание'] and i['Уровень'] == 'Проверить вручную'
               for i in audit_document(output, MISIS))


def test_table_opt_out_does_not_change_direct_cell_properties_or_structure():
    doc = sample(); source = save(doc)
    original = get_document(source).tables[0]._tbl.xml
    output, *_ = format_docx(source, MISIS, format_tables=False)
    assert get_document(output).tables[0]._tbl.xml == original
    assert table_errors(output)


def test_generic_and_custom_table_sizes_are_not_forced_to_misis():
    source = save(sample())
    for cfg, size, spacing in [(STANDARD, 12, 1.5),
                               (Profile(size=14, line_spacing=2), 14, 2),
                               (Profile(size=14, table_size=11, table_line_spacing=1.15), 11, 1.15)]:
        output, *_ = format_docx(source, cfg)
        p = get_document(output).tables[0].cell(0, 0).paragraphs[0]
        assert p.runs[0].font.size.pt == size
        assert p.paragraph_format.line_spacing == spacing


def test_profile_json_legacy_compatibility_and_misis_validation():
    assert profile_from_json(profile_to_json(MISIS)) == MISIS
    legacy = {key: value for key, value in STANDARD.summary().items()
              if key not in ('ruleset', 'table_size', 'table_line_spacing')}
    assert profile_from_json(json.dumps(legacy).encode()) == STANDARD
    for change in [{'table_size': True}, {'table_line_spacing': 99},
                   {'ruleset': 'unknown'}, {'table_size': 12}, {'size': 14}]:
        with pytest.raises(ValueError):
            profile_from_json(json.dumps({**MISIS.summary(), **change}).encode())


def test_audit_is_read_only():
    source = save(sample())
    before = bytes(source)
    assert audit_document(source, MISIS) == audit_document(source, MISIS)
    assert source == before


def test_preserves_images_math_links_lists_bookmarks_fields_headers_and_sections():
    doc = sample(); cell = doc.tables[0].cell(1, 0)
    p = cell.paragraphs[0]
    image = io.BytesIO(); Image.new('RGB', (2, 2), 'white').save(image, format='PNG'); image.seek(0)
    p.add_run().add_picture(image, width=Cm(.2))
    math = OxmlElement('m:oMath'); math.append(OxmlElement('m:r')); p._p.append(math)
    bookmark = OxmlElement('w:bookmarkStart'); bookmark.set(qn('w:id'), '42'); bookmark.set(qn('w:name'), 'KeepMe'); p._p.append(bookmark)
    bookmark_end = OxmlElement('w:bookmarkEnd'); bookmark_end.set(qn('w:id'), '42'); p._p.append(bookmark_end)
    field = OxmlElement('w:fldSimple'); field.set(qn('w:instr'), ' REF KeepMe '); p._p.append(field)
    link = OxmlElement('w:hyperlink'); link.set(qn('w:anchor'), 'KeepMe'); p._p.append(link)
    ref = OxmlElement('w:footnoteReference'); ref.set(qn('w:id'), '1'); p.add_run()._r.append(ref)
    cell.add_paragraph('Элемент списка', style='List Bullet')
    doc.sections[0].header.paragraphs[0].text = 'Существующая шапка'
    doc.sections[0].footer.paragraphs[0].text = 'Существующий колонтитул'
    doc.add_section()
    source = save(doc)
    # Attach a valid native Word footnotes part, with only synthetic text.
    with zipfile.ZipFile(io.BytesIO(source)) as original:
        parts = {name: original.read(name) for name in original.namelist()}
    parts['word/footnotes.xml'] = b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Keep footnote</w:t></w:r></w:p></w:footnote></w:footnotes>'
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(b'</Relationships>', b'<Relationship Id="rIdFootnotesTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'</Types>', b'<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
    payload = io.BytesIO()
    with zipfile.ZipFile(payload, 'w') as z:
        for name, value in parts.items():
            z.writestr(name, value)
    output, *_ = format_docx(payload.getvalue(), MISIS)
    fixed = get_document(output)
    xml = fixed.element.xml
    for tag in ['w:drawing', 'm:oMath', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:fldSimple',
                'w:hyperlink', 'w:footnoteReference']:
        assert fixed.element.xpath(f'.//{tag}')
    assert fixed.tables[0].cell(1, 0).paragraphs[1].style.name == 'List Bullet'
    assert len(fixed.sections) == 2
    assert fixed.sections[0].header.paragraphs[0].text == 'Существующая шапка'
    assert 'Существующий колонтитул' in '\n'.join(p.text for p in fixed.sections[0].footer.paragraphs)
    with zipfile.ZipFile(io.BytesIO(output)) as z:
        assert z.read('word/footnotes.xml') == parts['word/footnotes.xml']
        for name, value in parts.items():
            if name.startswith('word/media/'):
                assert z.read(name) == value
