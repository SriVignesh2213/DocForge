import logging
import re
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
    NUMBERING_RE = re.compile(r'^\s*(\d+(?:\.\d+)*)[\.\)]?\s+')
    CAPTION_RE = re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+)?(?:table|figure|fig)\b', re.IGNORECASE)
    MAJOR_HEADINGS = {
        "abstract",
        "introduction",
        "materials and methods",
        "materials & methods",
        "method",
        "methodology",
        "results",
        "results and discussion",
        "discussion",
        "conclusion",
        "conclusions",
        "acknowledgement",
        "acknowledgements",
        "references",
        "authors' note",
        "authors’ note",
    }

    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = Document(self.doc_path)
        self.sections = []
        self._parse()

    def _normalized_text(self, text):
        return self.NUMBERING_RE.sub("", text, count=1).strip().lower()

    def _paragraph_metrics(self, paragraph):
        max_size = 0
        is_bold = False
        is_italic = False
        for run in paragraph.runs:
            if run.text.strip():
                if run.bold:
                    is_bold = True
                if run.italic:
                    is_italic = True
                if run.font.size and run.font.size.pt > max_size:
                    max_size = run.font.size.pt
        return is_bold, is_italic, max_size

    def is_heading(self, p):
        if not isinstance(p, Paragraph):
            return False
        text = p.text.strip()
        if not text:
            return False

        style_name = (p.style.name or "").lower()
        if 'heading' in style_name:
            return True
        if self.CAPTION_RE.match(text) or text.lower().startswith("keywords:"):
            return False

        normalized = self._normalized_text(text)
        is_bold, is_italic, max_size = self._paragraph_metrics(p)
        word_count = len(text.split())

        if normalized in self.MAJOR_HEADINGS:
            return True
        if is_italic and word_count > 3:
            return False
        if self.NUMBERING_RE.match(text) and word_count <= 12:
            return True
        if is_bold and max_size >= 11 and word_count <= 12:
            return True

        letters = [char for char in normalized if char.isalpha()]
        if letters and sum(char.isupper() for char in text if char.isalpha()) >= max(3, int(len(letters) * 0.6)) and word_count <= 12:
            return True

        return False
    
    def get_heading_level(self, p):
        style_name = (p.style.name or "").lower()
        digits = ''.join(filter(str.isdigit, style_name))
        if digits:
            try:
                return 2 if int(digits) >= 2 else 1
            except Exception:
                return 1

        text = p.text.strip()
        numbering = self.NUMBERING_RE.match(text)
        if numbering:
            return 2 if '.' in numbering.group(1) else 1

        normalized = self._normalized_text(text)
        if normalized.startswith("sub heading") or normalized.startswith("subheading"):
            return 2
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
