import logging
import os
from docx import Document
from docxcompose.composer import Composer
from doc_utils import iter_block_items
from docx.text.paragraph import Paragraph
from docx.table import Table
from table_handler import TableHandler
from image_handler import ImageHandler

logger = logging.getLogger(__name__)

class Formatter:
    def __init__(self, template_path, input_path, output_path):
        self.template_path = template_path
        self.input_path = input_path
        self.output_path = output_path
        self.table_handler = TableHandler()
        self.image_handler = ImageHandler()

    def apply_styles_and_build(self, input_sections, mapping):
        input_doc = Document(self.input_path)
        
        # 1. Update styles in input document based on mapping
        for in_sec in input_sections:
            target_template = mapping.get(in_sec.title)
            if not target_template:
                continue
                
            for element in in_sec.elements:
                # Need to find the corresponding element in the loaded input_doc
                # For simplicity, docx compose keeps the objects in order
                pass
                
        # A more direct approach to modifying the input doc styles:
        current_mapped_template = None
        for block in iter_block_items(input_doc):
            if isinstance(block, Paragraph):
                style_name = block.style.name.lower()
                if 'heading' in style_name:
                    title = block.text.strip()
                    if title in mapping:
                        current_mapped_template = mapping[title]
                        try:
                            block.style = current_mapped_template.heading_style
                        except:
                            pass
                else:
                    if current_mapped_template:
                        try:
                            block.style = current_mapped_template.body_style
                        except:
                            pass
            elif isinstance(block, Table):
                self.table_handler.apply_template_table_style(block)
                
        self.image_handler.validate_images(input_doc)

        styled_input_path = "styled_temp.docx"
        input_doc.save(styled_input_path)

        # 2. Rebuild using Composer to inherit template master properties
        template_doc = Document(self.template_path)
        
        # Clear body of template doc to leave only styles, headers, footers
        body = template_doc.element.body
        for element in list(body.iterchildren()):
            if not element.tag.endswith('sectPr'):
                body.remove(element)
                
        # Merge input doc into the empty template envelope
        composer = Composer(template_doc)
        styled_doc = Document(styled_input_path)
        composer.append(styled_doc)
        
        # MANDATORY: Retain the exact template header and footer across all appended sections
        if len(composer.doc.sections) > 1:
            for section in composer.doc.sections[1:]:
                section.header.is_linked_to_previous = True
                section.footer.is_linked_to_previous = True
        
        composer.save(self.output_path)
        logger.info(f"Successfully generated formatted document {self.output_path}")
        
        if os.path.exists(styled_input_path):
            os.remove(styled_input_path)
