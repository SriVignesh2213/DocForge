import logging
from docx import Document

logger = logging.getLogger(__name__)

class StyleExtractor:
    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = Document(self.doc_path)
        self.styles = self._extract()

    def _extract(self):
        styles = {}
        for s in self.doc.styles:
            style_info = {
                'name': s.name,
                'type': s.type,
                'font_name': s.font.name if hasattr(s, 'font') and s.font else None,
                'font_size': s.font.size.pt if hasattr(s, 'font') and s.font and s.font.size else None,
            }
            styles[s.name] = style_info
        logger.info(f"Extracted {len(styles)} styles from template")
        return styles

    def get_style(self, name):
        return self.styles.get(name)
