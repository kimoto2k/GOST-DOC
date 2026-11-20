import io

import pymupdf
import pytest
from docx import Document

from gostdoc.preview import pdf_preview, reference_preview, renderer_path, render_docx


def source():
    doc = Document(); doc.add_paragraph('Synthetic original title. Не менять смысл.')
    out = io.BytesIO(); doc.save(out); return out.getvalue()


def test_pdf_reference_preview_is_real_and_pages_can_be_selected():
    with pymupdf.open() as doc:
        doc.new_page().insert_text((50,50),'Original reference')
        doc.new_page().insert_text((50,50),'Second page')
        content = doc.tobytes()
    first = reference_preview(content, is_pdf=True)
    second = pdf_preview(content, start_page=1)
    assert first['page_count'] == 2
    assert first['images'][0].startswith(b'\x89PNG')
    assert first['images'][0] != second['images'][0]
    with pytest.raises(ValueError): pdf_preview(content, start_page=2)


def test_missing_renderer_does_not_fake_a_preview(monkeypatch):
    monkeypatch.setenv('GOSTDOC_SOFFICE','/nonexistent/soffice')
    with pytest.raises(ValueError,match='не установлен'):
        render_docx(source())


def test_invalid_pdf_and_oversized_reference_fail_cleanly():
    with pytest.raises(ValueError): reference_preview(b'not pdf',is_pdf=True)
    with pytest.raises(ValueError): reference_preview(b'x'*(15*1024*1024+1),is_pdf=True)


@pytest.mark.skipif(not renderer_path(), reason='LibreOffice Writer not installed')
def test_real_libreoffice_docx_conversion_preserves_original_input():
    content=source(); original=bytes(content)
    preview=render_docx(content)
    assert preview['page_count'] == 1
    assert preview['pdf'].startswith(b'%PDF')
    assert preview['images'][0].startswith(b'\x89PNG')
    assert content == original
    with pymupdf.open(stream=preview['pdf'],filetype='pdf') as doc:
        assert 'Synthetic original title' in doc[0].get_text()


@pytest.mark.skipif(not renderer_path(), reason='LibreOffice Writer not installed')
def test_automatic_heading_numbers_are_visible_in_rendered_pages():
    from gostdoc.formatter import MISIS, format_docx
    doc=Document()
    for text,level in [('1 First',1),('1.1 Child',2),('2 Second',1)]:
        doc.add_heading(text,level)
    out=io.BytesIO();doc.save(out)
    result,*_=format_docx(out.getvalue(),MISIS)
    preview=render_docx(result)
    with pymupdf.open(stream=preview['pdf'],filetype='pdf') as rendered:
        text=' '.join(page.get_text() for page in rendered)
    import re
    for expected in [r'1\s+First',r'1\.1\s+Child',r'2\s+Second']:
        assert re.search(expected,text)
