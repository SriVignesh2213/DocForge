import logging
from docx import Document

logger = logging.getLogger(__name__)

class TableHandler:
    def __init__(self):
        pass

    def apply_template_table_style(self, table, default_style_name="Table Grid"):
        """ Applies a standard template style to matched tables if required """
        try:
            table.style = default_style_name
        except Exception as e:
            logger.warning(f"Failed to apply style {default_style_name} to table: {e}")
