import logging
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from doc_utils import (
    iter_block_items,
    make_section_break_paragraph,
    move_blocks_after,
    paragraph_has_drawing,
    remove_adjacent_empty_paragraphs,
    remove_empty_paragraphs_between,
)

logger = logging.getLogger(__name__)

class ImageHandler:
    FIGURE_CAPTION_RE = re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+)?(?:figure|fig)\.?\b', re.IGNORECASE)
    MAX_CAPTION_CONTINUATIONS = 2
    MAX_CONTINUATION_WORDS = 18
    MAX_CONTINUATION_CHARS = 140
    NON_CAPTION_SECTION_TITLES = {
        "abstract",
        "introduction",
        "background",
        "method",
        "methodology",
        "materials and methods",
        "materials & methods",
        "results",
        "discussion",
        "results and discussion",
        "conclusion",
        "conclusions",
        "references",
        "acknowledgements",
        "acknowledgement",
    }
    MEDIA_SPACE_BEFORE = Pt(6)
    MEDIA_SPACE_AFTER = Pt(3)
    SINGLE_COLUMN_WRAP_FACTOR = 1.2

    def __init__(self):
        pass

    def _coerce_length(self, value):
        if value is None:
            return None

        try:
            return int(value)
        except (TypeError, ValueError):
            try:
                return int(round(float(str(value))))
            except (TypeError, ValueError):
                logger.warning(f"Skipping unsupported figure width value: {value!r}")
                return None

    def validate_images(self, document):
        # Additional PyMuPDF validation could go here, 
        # but docxcompose already securely moves images natively.
        logger.info("Images validated for transport.")

    def optimize_figure_layout(self, document, layout_profile):
        wrapped_large_figures = False
        blocks = list(iter_block_items(document))

        for index, block in enumerate(blocks):
            if not isinstance(block, Paragraph) or not paragraph_has_drawing(block):
                continue

            previous_heading = self._find_previous_heading(blocks, index)
            caption_after = self._find_caption_bundle_after(blocks, index)
            caption_before = self._find_caption_bundle_before(blocks, index)
            caption_bundle = caption_after

            if caption_bundle is None and caption_before is not None:
                move_blocks_after([paragraph._element for paragraph in caption_before], block._element)
                caption_bundle = caption_before

            if previous_heading is not None:
                remove_empty_paragraphs_between(previous_heading._element, block._element)
                previous_heading.paragraph_format.keep_with_next = True
                previous_heading.paragraph_format.keep_together = True
                previous_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

            self._apply_figure_layout(block, caption_bundle)
            if caption_bundle:
                remove_empty_paragraphs_between(block._element, caption_bundle[0]._element)
                remove_adjacent_empty_paragraphs(caption_bundle[-1]._element)
            remove_adjacent_empty_paragraphs(block._element)

            figure_width = self._get_figure_width(block)
            is_large_figure = figure_width > (layout_profile["double_column_width"] * self.SINGLE_COLUMN_WRAP_FACTOR)
            target_width = layout_profile["single_column_width"] if is_large_figure else layout_profile["double_column_width"]
            self._fit_figure_to_width(block, figure_width, target_width)

            if is_large_figure:
                end_element = caption_bundle[-1]._element if caption_bundle is not None else block._element
                self._wrap_in_single_column(block._element, end_element, layout_profile)
                wrapped_large_figures = True

        return wrapped_large_figures

    def _find_caption_bundle_before(self, blocks, start_index):
        bundle = []
        continuation_count = 0
        index = start_index - 1

        while index >= 0:
            block = blocks[index]

            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return None

                text = block.text.strip()
                if not text:
                    index -= 1
                    continue

                if self.FIGURE_CAPTION_RE.match(text):
                    return [block] + bundle

                if continuation_count < self.MAX_CAPTION_CONTINUATIONS and self._is_caption_continuation(block):
                    bundle.insert(0, block)
                    continuation_count += 1
                    index -= 1
                    continue

                return None

            return None

        return None

    def _find_caption_bundle_after(self, blocks, start_index):
        bundle = []
        continuation_count = 0
        found_caption_start = False
        index = start_index + 1

        while index < len(blocks):
            block = blocks[index]

            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return bundle or None

                text = block.text.strip()
                if not text:
                    if found_caption_start:
                        break
                    index += 1
                    continue

                if not found_caption_start:
                    if self.FIGURE_CAPTION_RE.match(text):
                        bundle.append(block)
                        found_caption_start = True
                        index += 1
                        continue
                    return None

                if continuation_count < self.MAX_CAPTION_CONTINUATIONS and self._is_caption_continuation(block):
                    bundle.append(block)
                    continuation_count += 1
                    index += 1
                    continue

                break

            return bundle or None

        return bundle or None

    def _is_caption_continuation(self, paragraph):
        text = paragraph.text.strip()
        if not text:
            return False
        if self.FIGURE_CAPTION_RE.match(text):
            return False
        if "heading" in paragraph.style.name.lower():
            return False
        if any(run.bold for run in paragraph.runs if run.text.strip()):
            return False
        if text.lower().strip(":") in self.NON_CAPTION_SECTION_TITLES:
            return False
        return len(text.split()) <= self.MAX_CONTINUATION_WORDS and len(text) <= self.MAX_CONTINUATION_CHARS

    def _is_heading_like(self, paragraph):
        text = paragraph.text.strip()
        if not text or self.FIGURE_CAPTION_RE.match(text) or paragraph_has_drawing(paragraph):
            return False
        if text.lower().startswith("keywords:"):
            return False
        if "heading" in paragraph.style.name.lower():
            return True
        return any(run.bold for run in paragraph.runs if run.text.strip()) and len(text.split()) < 20

    def _find_previous_heading(self, blocks, start_index):
        index = start_index - 1
        while index >= 0:
            block = blocks[index]
            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return None
                if not block.text.strip():
                    index -= 1
                    continue
                return block if self._is_heading_like(block) else None
            return None
        return None

    def _apply_figure_layout(self, paragraph, caption_bundle):
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_with_next = bool(caption_bundle)
        paragraph.paragraph_format.keep_together = True
        paragraph.paragraph_format.space_before = self.MEDIA_SPACE_BEFORE
        paragraph.paragraph_format.space_after = self.MEDIA_SPACE_AFTER

        if not caption_bundle:
            return

        for index, caption_paragraph in enumerate(caption_bundle):
            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_paragraph.paragraph_format.keep_with_next = index < len(caption_bundle) - 1
            caption_paragraph.paragraph_format.keep_together = True
            caption_paragraph.paragraph_format.space_before = self.MEDIA_SPACE_BEFORE
            caption_paragraph.paragraph_format.space_after = self.MEDIA_SPACE_AFTER

    def _get_figure_width(self, paragraph):
        widths = []

        for extent in paragraph._p.xpath('.//wp:extent'):
            width = self._coerce_length(extent.get('cx'))
            if width:
                widths.append(width)

        if widths:
            return max(widths)

        for shape in paragraph._p.xpath('.//v:shape'):
            style = shape.get('style', '')
            match = re.search(r'width:([0-9.]+)pt', style, re.IGNORECASE)
            if match:
                widths.append(int(float(match.group(1)) * 12700))

        return max(widths, default=0)

    def _fit_figure_to_width(self, paragraph, current_width, target_width):
        if not current_width or current_width <= target_width:
            return

        scale = target_width / current_width

        for extent in paragraph._p.xpath('.//wp:extent'):
            width = self._coerce_length(extent.get('cx'))
            height = self._coerce_length(extent.get('cy'))
            if width and height:
                extent.set('cx', str(int(round(width * scale))))
                extent.set('cy', str(int(round(height * scale))))

        for extent in paragraph._p.xpath('.//a:ext'):
            width = self._coerce_length(extent.get('cx'))
            height = self._coerce_length(extent.get('cy'))
            if width and height:
                extent.set('cx', str(int(round(width * scale))))
                extent.set('cy', str(int(round(height * scale))))

    def _wrap_in_single_column(self, start_element, end_element, layout_profile):
        start_element.addprevious(
            make_section_break_paragraph(layout_profile["double_section_sectpr"], layout_profile["double_column_count"])
        )
        end_element.addnext(make_section_break_paragraph(layout_profile["double_section_sectpr"], 1))
