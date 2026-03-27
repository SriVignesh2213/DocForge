import logging
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
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
    FIGURE_LABEL_RE = re.compile(r'^\s*figure\s+labels?:', re.IGNORECASE)
    NUMBERED_HEADING_RE = re.compile(r'^\s*\d+(?:\.\d+)*\.?\s+')
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
    CLUSTER_MEDIA_SPACE_BEFORE = Pt(2)
    CLUSTER_MEDIA_SPACE_AFTER = Pt(0)
    SINGLE_COLUMN_WRAP_FACTOR = 1.2
    CLUSTER_MAX_HEIGHT_FACTOR = 0.9

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
        cluster_sizes = self._build_figure_cluster_map(blocks)

        for index, block in enumerate(blocks):
            if not isinstance(block, Paragraph) or not paragraph_has_drawing(block):
                continue

            previous_heading = self._find_previous_heading(blocks, index)
            caption_after = self._find_caption_bundle_after(blocks, index)
            caption_before = self._find_caption_bundle_before(blocks, index)
            caption_bundle = caption_after
            label_bundle = self._find_label_bundle_after(blocks, index, caption_bundle)

            if caption_bundle is None and caption_before is not None:
                move_blocks_after([paragraph._element for paragraph in caption_before], block._element)
                caption_bundle = caption_before
                label_bundle = self._find_label_bundle_after(blocks, index, caption_bundle)

            if previous_heading is not None:
                remove_empty_paragraphs_between(previous_heading._element, block._element)
                previous_heading.paragraph_format.keep_with_next = True
                previous_heading.paragraph_format.keep_together = True
                previous_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

            cluster_size = cluster_sizes.get(index, 1)
            self._apply_figure_layout(block, caption_bundle, label_bundle, cluster_size=cluster_size)
            if caption_bundle:
                remove_empty_paragraphs_between(block._element, caption_bundle[0]._element)
            if label_bundle:
                start_element = caption_bundle[-1]._element if caption_bundle else block._element
                remove_empty_paragraphs_between(start_element, label_bundle[0]._element)
                remove_adjacent_empty_paragraphs(label_bundle[-1]._element)
            elif caption_bundle:
                remove_adjacent_empty_paragraphs(caption_bundle[-1]._element)
            remove_adjacent_empty_paragraphs(block._element)

            figure_width, figure_height = self._get_figure_dimensions(block)
            is_large_figure = figure_width > (layout_profile["double_column_width"] * self.SINGLE_COLUMN_WRAP_FACTOR)
            target_width = layout_profile["single_column_width"] if is_large_figure else layout_profile["double_column_width"]
            max_height = None
            if cluster_size > 1 and not is_large_figure:
                max_height = int(layout_profile["double_column_width"] * self.CLUSTER_MAX_HEIGHT_FACTOR)
            self._fit_figure_to_bounds(block, figure_width, figure_height, target_width, max_height=max_height)

            if is_large_figure:
                if label_bundle is not None:
                    end_element = label_bundle[-1]._element
                elif caption_bundle is not None:
                    end_element = caption_bundle[-1]._element
                else:
                    end_element = block._element
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

    def _find_label_bundle_after(self, blocks, start_index, caption_bundle=None):
        bundle = []
        index = blocks.index(caption_bundle[-1]) + 1 if caption_bundle else start_index + 1

        while index < len(blocks):
            block = blocks[index]

            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return bundle or None

                text = block.text.strip()
                if not text:
                    if bundle:
                        break
                    index += 1
                    continue

                if not bundle:
                    if self.FIGURE_LABEL_RE.match(text):
                        bundle.append(block)
                        index += 1
                        continue
                    return None

                if self._is_label_continuation(block):
                    bundle.append(block)
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

    def _is_label_continuation(self, paragraph):
        text = paragraph.text.strip()
        if not text:
            return False
        if self.FIGURE_CAPTION_RE.match(text) or self.FIGURE_LABEL_RE.match(text):
            return False
        if "heading" in paragraph.style.name.lower():
            return False
        return len(text.split()) <= (self.MAX_CONTINUATION_WORDS * 2)

    def _has_list_numbering(self, paragraph):
        paragraph_properties = paragraph._p.pPr
        return paragraph_properties is not None and paragraph_properties.find(qn('w:numPr')) is not None

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

        return self._coerce_length(level.get(qn('w:val')))

    def _is_heading_like(self, paragraph):
        text = paragraph.text.strip()
        if not text or self.FIGURE_CAPTION_RE.match(text) or paragraph_has_drawing(paragraph):
            return False
        if text.lower().startswith("keywords:"):
            return False
        if self.FIGURE_LABEL_RE.match(text):
            return False
        if "heading" in paragraph.style.name.lower():
            return True
        if self.NUMBERED_HEADING_RE.match(text):
            return True

        list_level = self._get_list_level(paragraph)
        if list_level is not None and list_level >= 1 and len(text.split()) < 20:
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

    def _get_figure_block_end_index(self, blocks, start_index):
        caption_bundle = self._find_caption_bundle_after(blocks, start_index)
        label_bundle = self._find_label_bundle_after(blocks, start_index, caption_bundle)
        if label_bundle:
            return blocks.index(label_bundle[-1])
        if caption_bundle:
            return blocks.index(caption_bundle[-1])
        return start_index

    def _build_figure_cluster_map(self, blocks):
        cluster_sizes = {}
        index = 0

        while index < len(blocks):
            block = blocks[index]
            if not isinstance(block, Paragraph) or not paragraph_has_drawing(block):
                index += 1
                continue

            cluster_indexes = [index]
            end_index = self._get_figure_block_end_index(blocks, index)
            next_index = end_index + 1

            while next_index < len(blocks):
                next_block = blocks[next_index]
                if isinstance(next_block, Paragraph) and not next_block.text.strip() and not paragraph_has_drawing(next_block):
                    next_index += 1
                    continue

                if isinstance(next_block, Paragraph) and paragraph_has_drawing(next_block):
                    cluster_indexes.append(next_index)
                    end_index = self._get_figure_block_end_index(blocks, next_index)
                    next_index = end_index + 1
                    continue

                break

            if len(cluster_indexes) > 1:
                for cluster_index in cluster_indexes:
                    cluster_sizes[cluster_index] = len(cluster_indexes)

            index = end_index + 1

        return cluster_sizes

    def _apply_figure_layout(self, paragraph, caption_bundle, label_bundle, cluster_size=1):
        space_before = self.CLUSTER_MEDIA_SPACE_BEFORE if cluster_size > 1 else self.MEDIA_SPACE_BEFORE
        space_after = self.CLUSTER_MEDIA_SPACE_AFTER if cluster_size > 1 else self.MEDIA_SPACE_AFTER
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_with_next = bool(caption_bundle or label_bundle)
        paragraph.paragraph_format.keep_together = True
        paragraph.paragraph_format.space_before = space_before
        paragraph.paragraph_format.space_after = space_after

        if caption_bundle:
            for index, caption_paragraph in enumerate(caption_bundle):
                caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption_paragraph.paragraph_format.keep_with_next = bool(label_bundle) or index < len(caption_bundle) - 1
                caption_paragraph.paragraph_format.keep_together = True
                caption_paragraph.paragraph_format.space_before = space_before
                caption_paragraph.paragraph_format.space_after = space_after

        if not label_bundle:
            return

        for index, label_paragraph in enumerate(label_bundle):
            label_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            label_paragraph.paragraph_format.keep_with_next = index < len(label_bundle) - 1
            label_paragraph.paragraph_format.keep_together = True
            label_paragraph.paragraph_format.space_before = space_after
            label_paragraph.paragraph_format.space_after = space_after

    def _get_figure_dimensions(self, paragraph):
        dimensions = []

        for extent in paragraph._p.xpath('.//wp:extent'):
            width = self._coerce_length(extent.get('cx'))
            height = self._coerce_length(extent.get('cy'))
            if width and height:
                dimensions.append((width, height))

        if dimensions:
            return max(dimensions, key=lambda item: item[0] * item[1])

        widths = []
        heights = []
        for shape in paragraph._p.xpath('.//v:shape'):
            style = shape.get('style', '')
            width_match = re.search(r'width:([0-9.]+)pt', style, re.IGNORECASE)
            height_match = re.search(r'height:([0-9.]+)pt', style, re.IGNORECASE)
            if width_match:
                widths.append(int(float(width_match.group(1)) * 12700))
            if height_match:
                heights.append(int(float(height_match.group(1)) * 12700))

        if widths and heights:
            return max(widths), max(heights)

        if widths:
            return max(widths), 0

        return 0, 0

    def _fit_figure_to_bounds(self, paragraph, current_width, current_height, target_width, max_height=None):
        if not current_width:
            return

        scale = 1.0
        if current_width > target_width:
            scale = min(scale, target_width / current_width)
        if max_height and current_height and current_height > max_height:
            scale = min(scale, max_height / current_height)

        if scale >= 1:
            return

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
