from copy import deepcopy
from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif hasattr(parent, 'element') and hasattr(parent.element, 'body'):
        parent_elm = parent.element.body
    else:
        # Fallback for other elements that don't match exactly
        return

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def clone_section_properties(base_sectpr, column_count):
    sectpr = deepcopy(base_sectpr)

    section_type = sectpr.find(qn('w:type'))
    if section_type is None:
        section_type = OxmlElement('w:type')
        sectpr.insert(0, section_type)
    section_type.set(qn('w:val'), 'continuous')

    columns = sectpr.find(qn('w:cols'))
    if columns is None:
        columns = OxmlElement('w:cols')
        sectpr.append(columns)
    columns.set(qn('w:num'), str(column_count))

    return sectpr

def make_section_break_paragraph(base_sectpr, column_count):
    paragraph = OxmlElement('w:p')
    paragraph_properties = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    paragraph_properties.append(spacing)
    paragraph_properties.append(clone_section_properties(base_sectpr, column_count))
    paragraph.append(paragraph_properties)
    return paragraph

def move_block_before(source_element, target_element):
    source_element.getparent().remove(source_element)
    target_element.addprevious(source_element)

def move_block_after(source_element, target_element):
    source_element.getparent().remove(source_element)
    target_element.addnext(source_element)

def move_blocks_before(source_elements, target_element):
    elements = list(source_elements)
    for element in elements:
        element.getparent().remove(element)
    for element in elements:
        target_element.addprevious(element)

def move_blocks_after(source_elements, target_element):
    elements = list(source_elements)
    for element in elements:
        element.getparent().remove(element)

    anchor = target_element
    for element in elements:
        anchor.addnext(element)
        anchor = element

def paragraph_has_drawing(paragraph):
    return bool(paragraph._p.xpath('.//w:drawing')) or bool(paragraph._p.xpath('.//w:pict'))

def paragraph_has_section_break(paragraph):
    return bool(paragraph._p.xpath('.//w:sectPr'))

def _is_removable_empty_paragraph_element(element):
    if element is None or not isinstance(element, CT_P):
        return False
    if ''.join(element.itertext()).strip():
        return False
    if element.xpath('.//w:drawing') or element.xpath('.//w:pict') or element.xpath('.//w:sectPr'):
        return False
    return True

def remove_empty_paragraphs_between(start_element, end_element):
    current = start_element.getnext()
    while current is not None and current is not end_element:
        next_element = current.getnext()
        if _is_removable_empty_paragraph_element(current):
            current.getparent().remove(current)
        current = next_element

def remove_adjacent_empty_paragraphs(element):
    previous = element.getprevious()
    while _is_removable_empty_paragraph_element(previous):
        next_previous = previous.getprevious()
        previous.getparent().remove(previous)
        previous = next_previous

    following = element.getnext()
    while _is_removable_empty_paragraph_element(following):
        next_following = following.getnext()
        following.getparent().remove(following)
        following = next_following

def replace_body_section_properties(document, base_sectpr, column_count):
    body = document.element.body
    existing = body.sectPr
    new_sectpr = clone_section_properties(base_sectpr, column_count)

    if existing is not None:
        body.remove(existing)

    body.append(new_sectpr)
