import logging
from docx import Document
from doc_utils import iter_block_items
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)

class ParsedSection:
    def __init__(self, title, level):
        self.title = title
        self.level = level
        self.elements = []
        
class DocumentParser:
    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = Document(self.doc_path)
        self.sections = []
        self._parse()
        
    def is_heading(self, p):
        if not isinstance(p, Paragraph):
            return False
        style_name = p.style.name.lower()
        if 'heading' in style_name:
            return True
        return False
        
    def get_heading_level(self, p):
        style_name = p.style.name.lower()
        if 'heading' in style_name:
            try:
                # e.g. "heading 1" -> 1
                return int(''.join(filter(str.isdigit, style_name))) or 1
            except:
                return 1
        return 0

    def _parse(self):
        current_section = ParsedSection("Document Start", 0)
        self.sections.append(current_section)
        
        for block in iter_block_items(self.doc):
            if self.is_heading(block):
                level = self.get_heading_level(block)
                title = block.text.strip()
                if title:
                    current_section = ParsedSection(title, level)
                    self.sections.append(current_section)
            current_section.elements.append(block)
            
        logger.info(f"Parsed {len(self.sections)} sections from input document")
        
    def get_sections(self):
        return self.sections
