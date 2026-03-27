import logging
import re

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from doc_utils import (
    iter_block_items,
    make_section_break_paragraph,
    move_blocks_before,
    paragraph_has_drawing,
    remove_adjacent_empty_paragraphs,
    remove_empty_paragraphs_between,
)

logger = logging.getLogger(__name__)

class TableHandler:
    TABLE_CAPTION_RE = re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+)?table\b', re.IGNORECASE)
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
                logger.warning(f"Skipping unsupported table width value: {value!r}")
                return None

    def _get_grid_column_width(self, column):
        raw_width = column.get(qn('w:w'))
        if raw_width is not None:
            return self._coerce_length(raw_width)

        try:
            return self._coerce_length(column.w)
        except Exception as exc:
            logger.warning(f"Unable to read table grid width: {exc}")
            return None

    def _set_grid_column_width(self, column, width):
        column.set(qn('w:w'), str(int(round(width))))

    def _get_cell_width(self, cell):
        tc_pr = cell._tc.tcPr
        if tc_pr is not None and tc_pr.tcW is not None:
            raw_width = tc_pr.tcW.get(qn('w:w'))
            if raw_width is not None:
                return self._coerce_length(raw_width)

        try:
            return self._coerce_length(cell.width)
        except Exception as exc:
            logger.warning(f"Unable to read table cell width: {exc}")
            return None

    def apply_template_table_style(self, table, default_style_name="Table Grid"):
        """ Applies a standard template style to matched tables if required """
        try:
            table.style = default_style_name
        except Exception as e:
            logger.warning(f"Failed to apply style {default_style_name} to table: {e}")

    def optimize_table_layout(self, document, layout_profile):
        wrapped_large_tables = False
        self._remove_nearby_duplicate_tables(document)
        blocks = list(iter_block_items(document))

        for index, block in enumerate(blocks):
            if not isinstance(block, Table):
                continue

            previous_heading = self._find_previous_heading(blocks, index)
            self._normalize_table_position(block)

            caption_before = self._find_caption_bundle_before(blocks, index)
            caption_after = self._find_caption_bundle_after(blocks, index)
            caption_bundle = caption_before

            if caption_bundle is None and caption_after is not None:
                move_blocks_before([paragraph._element for paragraph in caption_after], block._element)
                caption_bundle = caption_after

            if previous_heading is not None:
                anchor_element = caption_bundle[0]._element if caption_bundle else block._element
                remove_empty_paragraphs_between(previous_heading._element, anchor_element)
                previous_heading.paragraph_format.keep_with_next = True
                previous_heading.paragraph_format.keep_together = True
                previous_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

            self._apply_table_layout(block, caption_bundle)
            if caption_bundle:
                remove_empty_paragraphs_between(caption_bundle[-1]._element, block._element)
                remove_adjacent_empty_paragraphs(caption_bundle[0]._element)
            remove_adjacent_empty_paragraphs(block._element)

            table_width = self._get_table_width(block)
            is_large_table = table_width > (layout_profile["double_column_width"] * self.SINGLE_COLUMN_WRAP_FACTOR)
            target_width = layout_profile["single_column_width"] if is_large_table else layout_profile["double_column_width"]
            self._fit_table_to_width(block, table_width, target_width)

            if is_large_table:
                start_element = caption_bundle[0]._element if caption_bundle is not None else block._element
                self._wrap_in_single_column(start_element, block._element, layout_profile)
                wrapped_large_tables = True

        return wrapped_large_tables

    def _find_caption_bundle_before(self, blocks, start_index):
        bundle = []
        continuation_count = 0
        index = start_index - 1

        while index >= 0:
            block = blocks[index]

            if isinstance(block, Table):
                return None
            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return None

                text = block.text.strip()
                if not text:
                    index -= 1
                    continue

                if self.TABLE_CAPTION_RE.match(text):
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

            if isinstance(block, Table):
                return bundle or None
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
                    if self.TABLE_CAPTION_RE.match(text):
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
        if self.TABLE_CAPTION_RE.match(text):
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
        if not text or self.TABLE_CAPTION_RE.match(text) or paragraph_has_drawing(paragraph):
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
            if isinstance(block, Table):
                return None
            if isinstance(block, Paragraph):
                if paragraph_has_drawing(block):
                    return None
                if not block.text.strip():
                    index -= 1
                    continue
                return block if self._is_heading_like(block) else None
            return None
        return None

    def _apply_table_layout(self, table, caption_bundle):
        try:
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
        except Exception as exc:
            logger.warning(f"Failed to center table: {exc}")

        if not caption_bundle:
            return

        for paragraph in caption_bundle:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.keep_with_next = True
            paragraph.paragraph_format.keep_together = True
            paragraph.paragraph_format.space_before = self.MEDIA_SPACE_BEFORE
            paragraph.paragraph_format.space_after = self.MEDIA_SPACE_AFTER

    def _normalize_table_position(self, table):
        table_properties = table._tbl.tblPr
        if table_properties is None:
            return

        floating_properties = table_properties.find(qn('w:tblpPr'))
        if floating_properties is not None:
            table_properties.remove(floating_properties)

    def _remove_nearby_duplicate_tables(self, document):
        changed = True
        while changed:
            changed = False
            blocks = list(iter_block_items(document))

            for index, block in enumerate(blocks):
                if not isinstance(block, Table):
                    continue

                caption_after = self._find_caption_bundle_after(blocks, index)
                if not caption_after:
                    continue

                next_table = self._find_next_table(blocks, index + 1)
                if next_table is None:
                    continue

                next_index, next_table_block = next_table
                caption_end_index = blocks.index(caption_after[-1])
                if next_index <= caption_end_index:
                    continue
                if next_index - caption_end_index > 1:
                    continue
                if self._find_caption_bundle_before(blocks, index):
                    continue
                if not self._tables_look_duplicate(block, next_table_block):
                    continue

                block._element.getparent().remove(block._element)
                changed = True
                break

    def _find_next_table(self, blocks, start_index):
        for index in range(start_index, len(blocks)):
            if isinstance(blocks[index], Table):
                return index, blocks[index]
            if isinstance(blocks[index], Paragraph) and blocks[index].text.strip():
                continue
        return None

    def _table_signature(self, table):
        signature = []
        for row in table.rows[:4]:
            for cell in row.cells[:6]:
                text = " ".join(cell.text.strip().lower().split())
                if text:
                    signature.append(text)
        return signature

    def _tables_look_duplicate(self, first_table, second_table):
        first_signature = self._table_signature(first_table)
        second_signature = self._table_signature(second_table)
        if not first_signature or not second_signature:
            return False

        first_set = set(first_signature)
        second_set = set(second_signature)
        overlap = len(first_set & second_set)
        required_overlap = max(4, int(min(len(first_set), len(second_set)) * 0.6))

        return (
            abs(len(first_table.columns) - len(second_table.columns)) <= 1
            and overlap >= required_overlap
        )

    def _get_table_width(self, table):
        row_widths = []
        for row in table.rows:
            cell_widths = []
            for cell in row.cells:
                width = self._get_cell_width(cell)
                if width:
                    cell_widths.append(width)
            if cell_widths:
                row_widths.append(sum(cell_widths))

        if row_widths:
            return max(row_widths)

        grid_widths = []
        for column in table._tbl.tblGrid.gridCol_lst:
            width = self._get_grid_column_width(column)
            if width:
                grid_widths.append(width)
        if grid_widths:
            return sum(grid_widths)

        return 0

    def _fit_table_to_width(self, table, current_width, target_width):
        if not current_width or current_width <= target_width:
            return

        scale = target_width / current_width

        for column in table._tbl.tblGrid.gridCol_lst:
            width = self._get_grid_column_width(column)
            if width:
                self._set_grid_column_width(column, width * scale)

        for row in table.rows:
            for cell in row.cells:
                width = self._get_cell_width(cell)
                if width:
                    cell.width = int(round(width * scale))

        try:
            table.autofit = False
        except Exception as exc:
            logger.warning(f"Failed to lock table width after scaling: {exc}")

    def _wrap_in_single_column(self, start_element, end_element, layout_profile):
        start_element.addprevious(
            make_section_break_paragraph(layout_profile["double_section_sectpr"], layout_profile["double_column_count"])
        )
        end_element.addnext(make_section_break_paragraph(layout_profile["double_section_sectpr"], 1))
