import logging
from docx import Document
from doc_utils import iter_block_items
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)

class TemplateSection:
    def __init__(self, title, level, style_name):
        self.title = title
        self.level = level
        self.heading_style = style_name
        self.body_style = "Normal"

class TemplateParser:
    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = Document(self.doc_path)
        self.sections = []
        self._parse()

    def is_heading(self, p):
        if not isinstance(p, Paragraph):
            return False
        return 'heading' in p.style.name.lower()
    
    def get_heading_level(self, p):
        style_name = p.style.name.lower()
        try:
            return int(''.join(filter(str.isdigit, style_name))) or 1
        except:
            return 1

    def _parse(self):
        for block in iter_block_items(self.doc):
            if self.is_heading(block):
                title = block.text.strip()
                if title:
                    level = self.get_heading_level(block)
                    style_name = block.style.name
                    sec = TemplateSection(title, level, style_name)
                    self.sections.append(sec)
            elif isinstance(block, Paragraph) and len(self.sections) > 0 and block.text.strip():
                if not 'heading' in block.style.name.lower():
                    self.sections[-1].body_style = block.style.name
                
        logger.info(f"Parsed {len(self.sections)} template sections.")
        
    def get_sections(self):
        return self.sections
