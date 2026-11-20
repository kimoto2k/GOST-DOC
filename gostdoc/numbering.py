"""Convert consistent heading labels to genuine multilevel Word numbering."""
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _node(name, value):
    element = OxmlElement('w:' + name)
    element.set(qn('w:val'), str(value))
    return element


def has_valid_heading_numbering(doc, paragraph, level):
    if not level or paragraph._p.pPr is None or paragraph._p.pPr.numPr is None:
        return False
    num = paragraph._p.pPr.numPr
    if num.numId is None or num.ilvl is None or num.ilvl.val != level - 1:
        return False
    root = doc.part.numbering_part.element
    numbers = root.xpath(f'./w:num[@w:numId="{num.numId.val}"]/w:abstractNumId')
    if not numbers:
        return False
    abstract = numbers[0].get(qn('w:val'))
    if abstract is None or not abstract.isdigit():
        return False
    levels = root.xpath(f'./w:abstractNum[@w:abstractNumId="{abstract}"]/w:lvl[@w:ilvl="{level-1}"]')
    if not levels:
        return False
    item = levels[0]
    for name, expected in [('numFmt','decimal'),('suff','space'),('lvlText','.'.join('%'+str(i+1) for i in range(level)))]:
        node = item.find(qn('w:'+name))
        if node is None or node.get(qn('w:val')) != expected:
            return False
    return True


def auto_number_headings(doc, guess=False, excluded=(), changed_positions=None):
    from .cover import inside_cover
    from .misis import heading_level, paragraph_rule
    plan = []
    counts = [0] * 9
    for p in doc.paragraphs:
        if inside_cover(p._p, excluded) or paragraph_rule(p, guess).role != 'Заголовок':
            continue
        level = heading_level(p, guess) - 1
        if level and not all(counts[:level]):
            return False  # Missing parent: do not invent a hierarchy.
        counts[level] += 1
        counts[level+1:] = [0] * (8-level)
        expected = '.'.join(map(str, counts[:level+1]))
        match = re.match(r'^(\d+(?:\.\d+){0,8})\s+', p.text)
        if match and (match[1] != expected or '\n' in match[0]):
            return False  # Renumbering could silently invalidate a textual link.
        plan.append((p, level, match.end() if match else 0))
    if not plan:
        return True
    root = doc.part.numbering_part.element
    abstract_id = 1 + max([int(n.get(qn('w:abstractNumId'))) for n in root.findall(qn('w:abstractNum'))] or [-1])
    num_id = 1 + max([int(n.get(qn('w:numId'))) for n in root.findall(qn('w:num'))] or [0])
    abstract = OxmlElement('w:abstractNum'); abstract.set(qn('w:abstractNumId'), str(abstract_id))
    abstract.append(_node('multiLevelType', 'multilevel'))
    for level in range(9):
        item = OxmlElement('w:lvl'); item.set(qn('w:ilvl'), str(level))
        item.extend([_node('start', 1), _node('numFmt', 'decimal'),
                     _node('lvlText', '.'.join('%' + str(n + 1) for n in range(level + 1))),
                     _node('suff', 'space'), _node('lvlJc', 'left')])
        ppr = OxmlElement('w:pPr')
        indent = OxmlElement('w:ind'); indent.set(qn('w:left'), '709'); indent.set(qn('w:firstLine'), '0')
        ppr.append(indent); item.append(ppr); abstract.append(item)
    # Abstract definitions precede num instances in the Word schema.
    existing_num = root.find(qn('w:num'))
    if existing_num is not None:
        existing_num.addprevious(abstract)
    else:
        root.append(abstract)
    number = OxmlElement('w:num'); number.set(qn('w:numId'), str(num_id))
    number.append(_node('abstractNumId', abstract_id))
    root.insert_element_before(number, 'w:numIdMacAtCleanup')
    for p, level, remove in plan:
        if remove and changed_positions is not None:
            changed_positions.add(list(doc.element.iter(qn('w:p'))).index(p._p))
        # Remove only the verified label, keeping runs, fields and bookmarks.
        for node in list(p._p.iter()):
            if not remove:
                break
            if next(node.iterancestors(qn('w:p')), None) is not p._p:
                continue
            if node.tag == qn('w:t'):
                value = node.text or ''
                amount = min(remove, len(value)); node.text = value[amount:]; remove -= amount
            elif node.tag == qn('w:tab'):
                node.getparent().remove(node); remove -= 1
        num = p._p.get_or_add_pPr().get_or_add_numPr()
        num.get_or_add_ilvl().val = level
        num.get_or_add_numId().val = num_id
    return True
