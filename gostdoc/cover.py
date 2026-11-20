"""Protect an explicitly bounded student cover from changes to Normal/Heading."""
import copy
import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def cover_elements(doc):
    prefix = []
    for node in doc.element.body:
        if node.tag == qn('w:sectPr'):
            break
        if node.tag == qn('w:sdt') and node.xpath('./w:sdtPr/w:alias[@w:val="Титульный лист"] | ./w:sdtPr/w:tag[@w:val="Титульный лист"]'):
            return [node]
        text = ''.join(node.xpath('.//w:t/text()')).strip()
        if node.tag == qn('w:p'):
            p = Paragraph(node, doc._body)
            if prefix and (text.upper() in ('ВВЕДЕНИЕ', 'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ')
                           or p.paragraph_format.page_break_before):
                break
        prefix.append(node)
        if node.tag == qn('w:p') and node.xpath('.//w:br[@w:type="page"] | ./w:pPr/w:sectPr'):
            break
        if len(prefix) > 100:
            return []
    combined = ' '.join(''.join(n.xpath('.//w:t/text()')) for n in prefix).upper()
    # A positive institutional/student signature AND a page/section boundary is
    # required. Never freeze an arbitrary first page of the report as a cover.
    ended = (prefix and (prefix[-1].xpath('.//w:br[@w:type="page"] | ./w:pPr/w:sectPr')
                         or prefix[-1].getnext() is not None and
                         prefix[-1].getnext().tag != qn('w:sectPr')))
    if ended and 'МИСИС' in combined and re.search(r'ГРУПП|ВЫПОЛНИЛ|СТУДЕНТ|ОБУЧАЮЩ', combined) and re.search(
            r'МИНИСТЕРСТВО|КАФЕДРА|УНИВЕРСИТЕТ', combined):
        return prefix
    return []


def cover_paragraphs(doc):
    return {p for node in cover_elements(doc) for p in ([node] if node.tag == qn('w:p') else node.xpath('.//w:p'))}


def protect_cover(doc):
    nodes = cover_elements(doc)
    if not nodes:
        return set()
    styles = {s.style_id: s.element for s in doc.styles}
    mapping = {}

    def clone(identifier):
        if identifier.startswith('GOSTCover_'):
            return identifier
        if identifier in mapping:
            return mapping[identifier]
        if identifier not in styles:
            return identifier
        safe = 'GOSTCover_' + identifier
        while safe in styles:
            safe = '_' + safe
        mapping[identifier] = safe
        element = copy.deepcopy(styles[identifier])
        element.set(qn('w:styleId'), safe)
        element.attrib.pop(qn('w:default'), None)
        name = element.find(qn('w:name'))
        if name is not None:
            name.set(qn('w:val'), safe)
        for name in ('basedOn', 'next', 'link'):
            link = element.find(qn(f'w:{name}'))
            if link is not None:
                link.set(qn('w:val'), clone(link.get(qn('w:val'))))
        doc.styles.element.append(element)
        return safe

    for node in nodes:
        paragraphs = [node] if node.tag == qn('w:p') else node.xpath('.//w:p')
        for element in paragraphs:
            p = Paragraph(element, doc._body)
            identifier = p.style.style_id
            p._p.get_or_add_pPr().get_or_add_pStyle().val = clone(identifier)
            for style in p._p.xpath('.//w:rPr/w:rStyle'):
                style.set(qn('w:val'), clone(style.get(qn('w:val'))))
    following = nodes[-1].getnext()
    if (following is not None and following.tag == qn('w:p')
            and not nodes[-1].xpath('.//w:br[@w:type="page"] | ./w:pPr/w:sectPr')):
        Paragraph(following, doc._body).paragraph_format.page_break_before = True
    return set(nodes)


def inside_cover(node, elements):
    return node in elements or any(parent in elements for parent in node.iterancestors())
