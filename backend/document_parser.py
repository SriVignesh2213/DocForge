import logging
import re
from docx import Document
from doc_utils import iter_block_items
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)

class ParsedSection:
    def __init__(self, title, level):
        self.title = title
        self.level = level
        self.elements = []
        
class DocumentParser:
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

    def _get_list_level(self, paragraph):
        paragraph_properties = paragraph._p.pPr
        if paragraph_properties is None:
            return None

        numbering_properties = paragraph_properties.find(qn('w:numPr'))
        if numbering_properties is None:
            return None

        level = numbering_properties.find(qn('w:ilvl'))
        if level is None:
            return 0

        try:
            return int(level.get(qn('w:val')))
        except (TypeError, ValueError):
            return 0

    def _get_list_num_id(self, paragraph):
        paragraph_properties = paragraph._p.pPr
        if paragraph_properties is None:
            return None

        numbering_properties = paragraph_properties.find(qn('w:numPr'))
        if numbering_properties is None:
            return None

        num_id = numbering_properties.find(qn('w:numId'))
        if num_id is None:
            return None

        try:
            return int(num_id.get(qn('w:val')))
        except (TypeError, ValueError):
            return None

    def _looks_like_sentence_item(self, text):
        stripped = self.NUMBERING_RE.sub("", text, count=1).strip()
        if not stripped:
            return False
        words = stripped.rstrip(':').split()
        return stripped.endswith(('.', ';', '?', '!')) or len(words) > 8

    def _looks_like_heading_label(self, text):
        stripped = self.NUMBERING_RE.sub("", text, count=1).strip()
        if not stripped:
            return False
        words = stripped.rstrip(':').split()
        return bool(words) and len(words) <= 10 and not self._looks_like_sentence_item(stripped)

    def _next_non_empty_paragraph(self, blocks, index):
        for next_index in range(index + 1, len(blocks)):
            block = blocks[next_index]
            if isinstance(block, Paragraph) and block.text.strip():
                return next_index, block
        return None, None

    def _previous_non_empty_paragraph(self, blocks, index):
        for previous_index in range(index - 1, -1, -1):
            block = blocks[previous_index]
            if isinstance(block, Paragraph) and block.text.strip():
                return previous_index, block
        return None, None

    def _is_contextual_list_heading(self, paragraph, blocks, index):
        list_level = self._get_list_level(paragraph)
        if list_level is None or not self._looks_like_heading_label(paragraph.text):
            return False

        current_num_id = self._get_list_num_id(paragraph)
        _, previous_paragraph = self._previous_non_empty_paragraph(blocks, index)
        _, next_paragraph = self._next_non_empty_paragraph(blocks, index)
        if next_paragraph is None:
            return False

        previous_level = self._get_list_level(previous_paragraph) if previous_paragraph is not None else None
        previous_num_id = self._get_list_num_id(previous_paragraph) if previous_paragraph is not None else None
        next_text = next_paragraph.text.strip()
        next_level = self._get_list_level(next_paragraph)
        next_num_id = self._get_list_num_id(next_paragraph)

        if previous_num_id == current_num_id and previous_level == list_level:
            if not (next_num_id == current_num_id and next_level is not None and next_level > list_level):
                return False

        if next_level is None:
            return not self.CAPTION_RE.match(next_text) and not next_text.lower().startswith("keywords:")

        if current_num_id is not None and next_num_id == current_num_id and next_level > list_level:
            return True

        if current_num_id is not None and next_num_id != current_num_id:
            return self._looks_like_heading_label(next_text) or self._looks_like_sentence_item(next_text)

        return False

    def _is_media_lead_in(self, text, blocks, index):
        stripped = self.NUMBERING_RE.sub("", text, count=1).strip()
        if not stripped.endswith(':'):
            return False

        _, next_paragraph = self._next_non_empty_paragraph(blocks, index)
        if next_paragraph is None:
            return False

        next_text = next_paragraph.text.strip()
        return bool(self.CAPTION_RE.match(next_text))
        
    def is_heading(self, p, blocks=None, index=None):
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
        if blocks is not None and index is not None and self._is_media_lead_in(text, blocks, index):
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
        if blocks is not None and index is not None and self._is_contextual_list_heading(p, blocks, index):
            return True
        if is_bold and max_size >= 11 and word_count <= 12:
            return True

        letters = [char for char in normalized if char.isalpha()]
        if letters and sum(char.isupper() for char in text if char.isalpha()) >= max(3, int(len(letters) * 0.6)) and word_count <= 12:
            return True

        return False
        
    def get_heading_level(self, p, blocks=None, index=None):
        style_name = (p.style.name or "").lower()
        if 'heading' in style_name:
            try:
                return int(''.join(filter(str.isdigit, style_name))) or 1
            except Exception:
                return 1

        text = p.text.strip()
        numbering = self.NUMBERING_RE.match(text)
        if numbering:
            return 2 if '.' in numbering.group(1) else 1

        normalized = self._normalized_text(text)
        if normalized in self.MAJOR_HEADINGS:
            return 1

        if blocks is not None and index is not None and self._is_contextual_list_heading(p, blocks, index):
            return 2
        if normalized.startswith("sub heading") or normalized.startswith("subheading"):
            return 2

        return 1
        
    def _parse(self):
        blocks = list(iter_block_items(self.doc))
        current_section = ParsedSection("Document Start", 0)
        self.sections.append(current_section)
        
        for index, block in enumerate(blocks):
            if self.is_heading(block, blocks, index):
                level = self.get_heading_level(block, blocks, index)
                title = block.text.strip()
                if title:
                    current_section = ParsedSection(title, level)
                    self.sections.append(current_section)
            current_section.elements.append(block)
            
        logger.info(f"Parsed {len(self.sections)} sections from input document")
        
    def get_sections(self):
        return self.sections
