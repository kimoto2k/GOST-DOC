"""Bounded DOCX → real PDF → page images; no made-up document mockups."""
import io
import os
import shutil
import signal
import subprocess
import tempfile
import zipfile
from pathlib import Path

import pymupdf
from lxml import etree

from .formatter import MAX_BYTES, validate_docx

MAX_PDF_BYTES = 50 * 1024 * 1024


def renderer_path():
    configured = os.environ.get('GOSTDOC_SOFFICE')
    if configured:
        path = Path(configured)
        return str(path) if path.is_file() and os.access(path, os.X_OK) else None
    return shutil.which('libreoffice') or shutil.which('soffice')


def pdf_preview(content, *, max_pages=1, start_page=0):
    max_pages = max(1, min(int(max_pages), 6))
    if len(content) > MAX_PDF_BYTES:
        raise ValueError('PDF для предпросмотра превышает 50 МБ.')
    try:
        with pymupdf.open(stream=content, filetype='pdf') as doc:
            if doc.needs_pass or not doc.page_count:
                raise ValueError('PDF защищён паролем или не содержит страниц.')
            if not 0 <= start_page < doc.page_count:
                raise ValueError('Страница PDF не существует.')
            images = []
            for page in list(doc.pages(start=start_page, stop=min(doc.page_count, start_page + max_pages))):
                scale = min(1200 / max(page.rect.width, 1), 1600 / max(page.rect.height, 1), 1.5)
                images.append(page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False).tobytes('png'))
            return {'images': images, 'page_count': doc.page_count, 'pdf': content}
    except (pymupdf.FileDataError, RuntimeError) as exc:
        raise ValueError('Не удалось открыть PDF предпросмотра.') from exc


def render_docx(content, *, max_pages=1):
    validate_docx(content)
    executable = renderer_path()
    if not executable:
        raise ValueError('Предпросмотр недоступен: на сервере не установлен LibreOffice Writer.')
    with zipfile.ZipFile(io.BytesIO(content)) as package:
        for name in package.namelist():
            if name.endswith('.rels'):
                root = etree.fromstring(package.read(name), etree.XMLParser(resolve_entities=False, no_network=True))
                for rel in root:
                    if rel.get('TargetMode') == 'External' and not rel.get('Type', '').endswith('/hyperlink'):
                        raise ValueError('Предпросмотр отключён: DOCX ссылается на внешние объекты. Файл не изменён.')
    with tempfile.TemporaryDirectory(prefix='gostdoc-preview-') as directory:
        root = Path(directory)
        incoming = root / 'report.docx'
        incoming.write_bytes(content)
        profile = root / 'office-profile'
        user = profile / 'user'; user.mkdir(parents=True)
        (user / 'registrymodifications.xcu').write_text(
            '<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
            '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>'
            '<item oor:path="/org.openoffice.Office.Writer/Content/Update">'
            '<prop oor:name="Link" oor:op="fuse"><value>0</value></prop></item></oor:items>')
        command = [executable, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--norestore',
                   '--nodefault', '--nolockcheck', '--convert-to', 'pdf:writer_pdf_Export',
                   '--outdir', str(root), str(incoming)]
        process = subprocess.Popen(command, cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   start_new_session=os.name != 'nt')
        try:
            process.communicate(timeout=45)
        except subprocess.TimeoutExpired as exc:
            if os.name != 'nt':
                os.killpg(process.pid, signal.SIGKILL)
            else:
                process.kill()
            process.communicate()
            raise ValueError('Предпросмотр остановлен: превышено время рендера 45 секунд.') from exc
        output = root / 'report.pdf'
        if process.returncode or not output.is_file():
            raise ValueError('LibreOffice не смог построить страницы этого документа. Исходник сохранён.')
        if output.stat().st_size > MAX_PDF_BYTES:
            raise ValueError('Результат рендера превышает 50 МБ.')
        return pdf_preview(output.read_bytes(), max_pages=max_pages)


def reference_preview(content, *, is_pdf=False):
    if len(content) > MAX_BYTES:
        raise ValueError('Эталон превышает допустимые 15 МБ.')
    return pdf_preview(content) if is_pdf else render_docx(content)
