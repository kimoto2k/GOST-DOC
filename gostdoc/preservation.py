"""Fail closed if a formatting operation loses document content or objects."""
import hashlib
import io
import re
import zipfile
import posixpath
from collections import Counter

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _text(value):
    value = re.sub(r'\s+', ' ', value).strip()
    if value.rstrip('.').upper() in ('ОГЛАВЛЕНИЕ', 'СОДЕРЖАНИЕ'):
        return 'СОДЕРЖАНИЕ'
    if re.fullmatch(r'ПРИЛОЖЕНИЕ(?:\s+[А-ЯЁ])?', value.rstrip('.').upper()) or value.rstrip('.').upper() in ('ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'ПРИЛОЖЕНИЕ',
                                    'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', 'СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ'):
        return value.rstrip('.').upper()
    return re.sub(r'^(Рисунок\s+\S+\s*)[—-]', r'\1–', value)


def _signature(element):
    return (element.tag, tuple(sorted(element.attrib.items())), element.text or '',
            tuple(_signature(child) for child in element))


def _manifest(data, *, heading_positions=()):
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    result = {'stories': {}, 'math': Counter(), 'media': Counter(), 'references': Counter(), 'tables': Counter(), 'relationships': Counter()}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in z.namelist():
            raw = z.read(name)
            if name.startswith('word/media/'):
                result['media'][hashlib.sha256(raw).hexdigest()] += 1
            if name.endswith('.rels'):
                relations = etree.fromstring(raw, parser)
                owner_folder = posixpath.dirname(posixpath.dirname(name))
                for relation in relations:
                    target = relation.get('Target', '')
                    if relation.get('TargetMode') != 'External':
                        target = posixpath.normpath(posixpath.join(owner_folder, target.lstrip('/')) if not target.startswith('/') else target[1:])
                    result['relationships'][(name, relation.get('Id'), relation.get('Type'), target, relation.get('TargetMode',''))] += 1
            if not name.startswith('word/') or not name.endswith('.xml'):
                continue
            root = etree.fromstring(raw, parser)
            texts = []
            for index, p in enumerate(root.iter(f'{{{W}}}p')):
                text = ''.join(node.text or '' if node.tag == f'{{{W}}}t' else '\t'
                               for node in p.iter() if node.tag in (f'{{{W}}}t', f'{{{W}}}tab'))
                if name == 'word/document.xml' and index in heading_positions:
                    text = re.sub(r'^\d+(?:\.\d+){0,8}\s+', '', text)
                texts.append(_text(text))
            result['stories'][name] = [value for value in texts if value]
            for math in root.iter(f'{{{M}}}oMath'):
                result['math'][_signature(math)] += 1
            for tag in ('bookmarkStart', 'bookmarkEnd', 'footnoteReference', 'endnoteReference', 'hyperlink', 'fldSimple'):
                for node in root.iter(f'{{{W}}}{tag}'):
                    result['references'][(tag, tuple(sorted(node.attrib.items())))] += 1
            for node in root.iter(f'{{{W}}}instrText'):
                result['references'][('field', (node.text or '').strip())] += 1
            for table in root.iter(f'{{{W}}}tbl'):
                rows = []
                for row in table.findall(f'{{{W}}}tr'):
                    rows.append(tuple((cell.find(f'{{{W}}}tcPr/{{{W}}}gridSpan').get(f'{{{W}}}val')
                                       if cell.find(f'{{{W}}}tcPr/{{{W}}}gridSpan') is not None else '1',
                                       cell.find(f'{{{W}}}tcPr/{{{W}}}vMerge').get(f'{{{W}}}val', 'continue')
                                       if cell.find(f'{{{W}}}tcPr/{{{W}}}vMerge') is not None else '')
                                      for cell in row.findall(f'{{{W}}}tc')))
                result['tables'][tuple(rows)] += 1
    return result


def assert_preserved(source, result, *, heading_positions=()):
    before = _manifest(source, heading_positions=heading_positions)
    after = _manifest(result)
    for name, paragraphs in before['stories'].items():
        cursor = iter(after['stories'].get(name, []))
        if not all(any(actual == expected for actual in cursor) for expected in paragraphs):
            raise ValueError('Результат заблокирован: изменён или утрачен исходный текст. Исходник сохранён.')
    for kind in ('math', 'media', 'references', 'tables', 'relationships'):
        if before[kind] - after[kind]:
            raise ValueError(f'Результат заблокирован: проверка сохранности объектов ({kind}) не пройдена.')
