import csv
import io
import json
import zipfile

import pytest
from docx import Document

from gostdoc.formatter import Profile, STANDARD, format_docx
from gostdoc.workflow import (
    audit_counts,
    audit_csv,
    audit_document,
    demo_docx,
    inspect_document,
    profile_from_json,
    profile_to_json,
    safe_docx_filename,
    zip_results,
)


def make_doc():
    doc = Document()
    doc.add_heading("ВВЕДЕНИЕ", level=1)
    doc.add_paragraph("Текст с ошибками форматирования. " * 3)
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "Таблица"
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def test_document_summary_and_audit():
    source = demo_docx()
    info = inspect_document(source)
    assert info["paragraphs"] >= 3
    assert info["words"] > 10
    issues = audit_document(source, STANDARD)
    assert issues
    assert audit_counts(issues)["errors"] >= 1
    assert all(set(item) == {"Уровень", "Категория", "Место", "Замечание"} for item in issues)


def test_csv_has_bom_and_semicolon_header():
    report = audit_csv([
        {"Уровень": "Ошибка", "Категория": "Страница", "Место": "Раздел 1",
         "Замечание": "Поля не соответствуют требованиям."}
    ])
    assert report.startswith(b"\xef\xbb\xbf")
    text = report.decode("utf-8-sig")
    assert "Уровень;Категория;Место;Замечание" in text
    assert "Поля не соответствуют" in text


def test_profile_json_round_trip():
    profile = Profile(font="Arial", size=13, right_cm=2.2)
    restored = profile_from_json(profile_to_json(profile))
    assert restored == profile
    assert json.loads(profile_to_json(profile))["size"] == 13


@pytest.mark.parametrize("payload", [
    b'{}',
    b'not json',
    b'{"font":"__import__(1)"}',
    b'x' * (16 * 1024 + 1),
])
def test_invalid_profile_rejected(payload):
    with pytest.raises(ValueError):
        profile_from_json(payload)


def test_zip_export_renames_duplicates_and_blocks_paths():
    binary = make_doc()
    output = zip_results([
        ("../../report.docx", binary),
        ("report.docx", binary),
    ], report=audit_csv([]))
    with zipfile.ZipFile(io.BytesIO(output)) as z:
        names = z.namelist()
        assert len(names) == 3
        assert "report_GOST.docx" in names
        assert "report_GOST_2.docx" in names
        assert "audit_report.csv" in names
        assert all(".." not in name and "/" not in name for name in names)


def test_name_sanitizer():
    assert safe_docx_filename(r"C:\documents\lab.docx") == "lab_GOST.docx"


def test_document_after_format_has_lower_error_count():
    data = demo_docx()
    before = audit_counts(audit_document(data, STANDARD))["errors"]
    output, *_ = format_docx(data, STANDARD)
    after = audit_counts(audit_document(output, STANDARD))["errors"]
    assert before > 0
    assert after <= before
