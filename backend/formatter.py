import logging
import os
from docx import Document
from docxcompose.composer import Composer
from doc_utils import iter_block_items
from docx.text.paragraph import Paragraph
from docx.table import Table
import re
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
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
                
        # -----------------------------
        # -----------------------------
        # EXACT 13 SOP FORMATTING RULES (SAFE MODE)
        # -----------------------------
        h1_num = 0
        h2_num = 0
        stop_num = False
        in_title_zone = True

        for p in [b for b in iter_block_items(input_doc) if isinstance(b, Paragraph)]:
            text = p.text.strip()
            text_lower = text.lower()
            if not text: continue
            
            # Rule 13 / Citation safe removal inside runs
            for r in p.runs:
                if r.text:
                    r.text = re.sub(r'\[\d+\]', '', r.text)

            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            max_size = 0
            is_bold = False
            for r in p.runs:
                if r.bold: is_bold = True
                if r.font.size and r.font.size.pt > max_size:
                    max_size = r.font.size.pt
                    
            is_heading_candidate = ('heading' in p.style.name.lower()) or (is_bold and 10 < max_size <= 14 and len(text.split()) < 20)

            # Rule 11: Title (16 pt, Bold, TNR, Center)
            if in_title_zone and (max_size >= 14 or p.style.name.startswith('Title')):
                in_title_zone = False
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(16)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0, 0, 0)
                continue
                
            # Rule 12: Author Info (Italic)
            if not in_title_zone and h1_num == 0 and not is_heading_candidate and len(text.split()) < 25:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(12)
                    r.font.italic = True
                    r.font.color.rgb = RGBColor(0, 0, 0)
                continue
                
            # Rules 5, 6, 7, 10: Headings
            if is_heading_candidate and len(text.split()) < 20:
                if "conclusion" in text_lower or "references" in text_lower:
                    stop_num = True
                    
                level = 2 if ('2' in p.style.name or (text.startswith('1.') and text.count('.') > 1)) else 1
                
                prefix = ""
                if not stop_num:
                    if level == 1:
                        h1_num += 1
                        h2_num = 0
                        prefix = f"{h1_num}. "
                    else:
                        h2_num += 1
                        prefix = f"{h1_num}.{h2_num}. "
                        
                # Soft strip preceding numbers in first run
                if p.runs and prefix:
                    p.runs[0].text = re.sub(r'^[\d\.]+\s*', '', p.runs[0].text.lstrip())
                    p.runs[0].text = prefix + p.runs[0].text
                
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(12)
                    r.font.bold = True
                    r.font.italic = False
                    r.font.color.rgb = RGBColor(0, 0, 255) # Rule 5: Blue
                continue
                
            # Standard Body application (Rule 1, 3, 4)
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)
                r.font.bold = False
                r.font.italic = False
                r.font.color.rgb = RGBColor(0, 0, 0)
                
        self.image_handler.validate_images(input_doc)
        
        styled_input_path = "styled_temp.docx"
        input_doc.save(styled_input_path)

        # Rebuild using Composer to inherit template master properties
        template_doc = Document(self.template_path)
        
        # Merge input doc into the empty template envelope natively
        
        # 3. Flawlessly strip template dummy bytes without killing its multi-section layout logic (where the header logos live!)
        for p in template_doc.paragraphs:
            if p._element.xpath('.//w:sectPr'):
                p.clear() # DO NOT delete node, it anchors an official layout section break!
            else:
                p._element.getparent().remove(p._element) # Safe to permanently delete dummy placeholder
                
        for t in template_doc.tables:
            t._element.getparent().remove(t._element) # Delete dummy tables
            
        composer = Composer(template_doc)
        styled_doc = Document(styled_input_path)
        composer.append(styled_doc)
        
        composer.save(self.output_path)
        logger.info(f"Successfully generated formatted document {self.output_path}")
        
        if os.path.exists(styled_input_path):
            os.remove(styled_input_path)
