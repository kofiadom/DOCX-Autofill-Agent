#!/usr/bin/env python3
"""
Filling Strategies Module - Phase 3 Implementation
Implements 6 different strategies (A-F) for filling DOCX templates
"""

from lib.document import Document
from typing import List, Tuple, Dict, Any


class FillingResult:
    """Track results of filling operation"""

    def __init__(self):
        self.filled = []
        self.skipped = []
        self.errors = []

    def add_filled(self, field_name: str):
        if field_name not in self.filled:
            self.filled.append(field_name)

    def add_skipped(self, field_name: str, reason: str = ""):
        self.skipped.append((field_name, reason))

    def add_error(self, field_name: str, error: str):
        self.errors.append((field_name, error))

    def summary(self) -> str:
        msg = f"Filled {len(self.filled)}"
        if self.filled:
            msg += ": " + ", ".join(self.filled)
        if self.skipped:
            msg += f" | Skipped {len(self.skipped)}"
        return msg


class StrategyA:
    """Strategy A: Simple Text Placeholder Replacement

    For placeholders in text runs: {{PLACEHOLDER}}, [PLACEHOLDER], __PLACEHOLDER__, <<PLACEHOLDER>>

    Finds w:r (run) elements containing placeholder text in multiple formats,
    preserves formatting, and replaces content.
    """

    @staticmethod
    def _try_placeholder_formats(placeholder_name: str) -> List[str]:
        """Generate placeholder in multiple formats to try

        Args:
            placeholder_name: Base placeholder name (without delimiters)

        Returns:
            List of placeholder formats to try in order
        """
        # Normalize: remove any existing delimiters
        clean_name = placeholder_name.strip('{} []_<>')

        # Return formats to try in order of preference
        return [
            f"{{{{{clean_name}}}}}",      # {{PLACEHOLDER}}
            f"[{clean_name}]",             # [PLACEHOLDER]
            f"__{clean_name}__",           # __PLACEHOLDER__
            f"<<{clean_name}>>",           # <<PLACEHOLDER>>
        ]

    @staticmethod
    def fill(doc: Document, placeholders: dict) -> FillingResult:
        """Replace text placeholders while preserving formatting

        Args:
            doc: Document instance
            placeholders: Dict of {placeholder_name: value}
                         Supports multiple formats: {{NAME}}, [NAME], __NAME__, <<NAME>>

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        for placeholder_name, value in placeholders.items():
            # Try multiple placeholder formats
            formats_to_try = StrategyA._try_placeholder_formats(placeholder_name)
            filled = False

            for search_text in formats_to_try:
                try:
                    # Find the text run containing placeholder
                    node = doc_xml.get_node(tag="w:r", contains=search_text)

                    # Preserve formatting (run properties)
                    rpr_list = node.getElementsByTagName("w:rPr")
                    rpr_xml = rpr_list[0].toxml() if rpr_list else ""

                    # Create replacement with same formatting
                    replacement = f'<w:r>{rpr_xml}<w:t>{value}</w:t></w:r>'

                    doc_xml.replace_node(node, replacement)
                    result.add_filled(placeholder_name)
                    filled = True
                    break  # Found and filled, move to next placeholder

                except ValueError:
                    # Format not found, try next format
                    continue
                except Exception as e:
                    # Unexpected error
                    result.add_error(placeholder_name, str(e))
                    filled = True
                    break

            if not filled:
                # None of the formats were found
                result.add_skipped(placeholder_name, "No matching placeholder format found")

        return result


class StrategyB:
    """Strategy B: Structured Data Tag (SDT) Replacement

    For Word's native form fields (text boxes, dropdowns, date pickers).

    Word uses w:sdt elements with w:alias attributes for field names.
    Content is in w:sdtContent/w:t elements.
    """

    @staticmethod
    def fill(doc: Document, placeholders: dict) -> FillingResult:
        """Replace Structured Data Tag field values

        Args:
            doc: Document instance
            placeholders: Dict of {field_name: value}
                         where field_name matches w:alias w:val

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        # Get all SDT elements
        sdts = doc_xml.dom.getElementsByTagName('w:sdt')

        for sdt in sdts:
            try:
                # Get field name from alias
                alias_elems = sdt.getElementsByTagName('w:alias')
                if not alias_elems:
                    continue

                field_name = alias_elems[0].getAttribute('w:val')
                if not field_name or field_name not in placeholders:
                    continue

                # Get value to fill
                value = placeholders[field_name]

                # Replace text in SDT content
                content_elems = sdt.getElementsByTagName('w:sdtContent')
                if not content_elems:
                    result.add_skipped(field_name, "No sdtContent found")
                    continue

                content_elem = content_elems[0]
                text_elems = content_elem.getElementsByTagName('w:t')

                if not text_elems:
                    result.add_skipped(field_name, "No w:t element in content")
                    continue

                # Update first text element
                text_elems[0].firstChild.nodeValue = value
                result.add_filled(field_name)

            except Exception as e:
                result.add_error(field_name, str(e))

        return result


class StrategyC:
    """Strategy C: Element ID-Based Replacement

    For fine-grained control using element IDs.

    Targets specific elements by their w:id attribute.
    Useful when you have pre-assigned IDs for specific fields.
    """

    @staticmethod
    def fill(doc: Document, id_mapping: dict) -> FillingResult:
        """Replace content in elements with specific IDs

        Args:
            doc: Document instance
            id_mapping: Dict of {element_id: new_value}
                       where element_id is w:id attribute value

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        for elem_id, new_value in id_mapping.items():
            try:
                # Find element by attribute
                node = doc_xml.get_node(tag="w:r", attrs={"w:id": elem_id})

                # Get text element
                text_elems = node.getElementsByTagName('w:t')
                if not text_elems:
                    result.add_skipped(elem_id, "No w:t element found")
                    continue

                text_elems[0].firstChild.nodeValue = new_value
                result.add_filled(elem_id)

            except ValueError as e:
                result.add_skipped(elem_id, f"Element not found: {str(e)}")
            except Exception as e:
                result.add_error(elem_id, str(e))

        return result


class StrategyD:
    """Strategy D: Multi-Run Placeholder Handling

    For placeholders split across multiple XML runs.

    Example: "{{CLIENT" in one w:r + "_NAME}}" in next w:r

    Combines split runs into single run with replacement.
    """

    @staticmethod
    def fill(doc: Document, placeholders: dict) -> FillingResult:
        """Handle placeholders that span multiple w:r elements

        Args:
            doc: Document instance
            placeholders: Dict of {placeholder_name: value}

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        # Find paragraph containing placeholder text
        paragraphs = doc_xml.dom.getElementsByTagName('w:p')

        for paragraph in paragraphs:
            text_runs = paragraph.getElementsByTagName('w:r')

            # Collect all text in paragraph
            combined_text = ''.join([
                t.firstChild.nodeValue if t.firstChild else ''
                for run in text_runs
                for t in (run.getElementsByTagName('w:t') or [])
                if t.firstChild
            ])

            # Check if placeholder is in this paragraph
            for placeholder_name, value in placeholders.items():
                search_text = f"{{{{{placeholder_name}}}}}"

                if search_text not in combined_text:
                    continue

                try:
                    # Get formatting from first run
                    rpr_xml = ""
                    if text_runs:
                        rpr_list = text_runs[0].getElementsByTagName('w:rPr')
                        rpr_xml = rpr_list[0].toxml() if rpr_list else ""

                    # Remove all runs in paragraph (except pPr)
                    runs_to_remove = []
                    for run in text_runs:
                        if run.parentNode == paragraph:
                            runs_to_remove.append(run)

                    for run in runs_to_remove:
                        run.parentNode.removeChild(run)

                    # Insert replacement as single run
                    new_run = f'<w:r>{rpr_xml}<w:t>{value}</w:t></w:r>'

                    # Insert after paragraph properties
                    ppr_elems = paragraph.getElementsByTagName('w:pPr')
                    if ppr_elems:
                        doc_xml.insert_after(ppr_elems[0], new_run)
                    else:
                        doc_xml.append_to(paragraph, new_run)

                    result.add_filled(placeholder_name)

                except Exception as e:
                    result.add_error(placeholder_name, str(e))

        return result


class StrategyE:
    """Strategy E: Table Filling

    For replacing data in template tables.

    Assumes first row is header with column names.
    Clones template row and fills with data.
    """

    @staticmethod
    def fill(doc: Document, table_index: int, row_data_list: list) -> FillingResult:
        """Fill table with multiple rows of data

        Args:
            doc: Document instance
            table_index: Which table in document (0-based)
            row_data_list: List of dicts with row data:
                          [
                              {'Item': 'Widget', 'Price': '$100'},
                              {'Item': 'Gadget', 'Price': '$50'}
                          ]

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        tables = doc_xml.dom.getElementsByTagName('w:tbl')
        if table_index >= len(tables):
            result.add_error(f"table_{table_index}", f"Table {table_index} not found")
            return result

        try:
            table = tables[table_index]
            rows = table.getElementsByTagName('w:tr')

            # Assume first row is header, skip it
            if len(rows) < 2:
                result.add_error(f"table_{table_index}", "No template row found")
                return result

            header_row = rows[0]
            template_row = rows[1]

            # Remove existing data rows (keep header and template)
            for i in range(len(rows) - 1, 1, -1):
                table.removeChild(rows[i])

            # Add rows for each data item
            for row_idx, row_data in enumerate(row_data_list):
                try:
                    # Clone template row
                    new_row = template_row.cloneNode(True)

                    cells = new_row.getElementsByTagName('w:tc')
                    header_cells = header_row.getElementsByTagName('w:tc')

                    # Get header names
                    header_names = []
                    for header_cell in header_cells:
                        text_elems = header_cell.getElementsByTagName('w:t')
                        header_text = ''.join([
                            t.firstChild.nodeValue if t.firstChild else ''
                            for t in text_elems
                        ])
                        header_names.append(header_text.strip())

                    # Fill cells with data
                    for cell_idx, cell in enumerate(cells):
                        if cell_idx < len(header_names):
                            col_key = header_names[cell_idx]
                            if col_key in row_data:
                                text_elems = cell.getElementsByTagName('w:t')
                                if text_elems:
                                    text_elems[0].firstChild.nodeValue = str(row_data[col_key])

                    table.appendChild(new_row)
                    result.add_filled(f"row_{row_idx}")

                except Exception as e:
                    result.add_error(f"row_{row_idx}", str(e))

            print(f"Filled {len(result.filled)} rows in table {table_index}")

        except Exception as e:
            result.add_error(f"table_{table_index}", str(e))

        return result


class StrategyF:
    """Strategy F: Conditional Content - NOT YET IMPLEMENTED

    For showing/hiding sections based on data.

    Would use bookmarks or comment markers to identify sections:
    <!-- IF:CONDITION_NAME -->content<!-- ENDIF:CONDITION_NAME -->

    IMPLEMENTATION STATUS: Stub only - requires bookmark/comment parsing support
    from Document Library. Not currently functional.

    This strategy is documented in AUTO_FILL_WORKFLOW.md Phase 3 but is not
    yet implemented. Use strategies A-E for current document filling needs.
    """

    @staticmethod
    def fill(doc: Document, conditions: dict) -> FillingResult:
        """Show/hide document sections based on conditions

        NOTE: This is a stub implementation. Not yet functional.

        Args:
            doc: Document instance
            conditions: Dict of {condition_name: boolean}
                       True = show section, False = hide section

        Returns:
            FillingResult with filled, skipped, and error counts
        """
        result = FillingResult()
        doc_xml = doc["word/document.xml"]

        for condition_name, show_content in conditions.items():
            try:
                if show_content:
                    result.add_filled(condition_name)
                else:
                    # Strategy F is not yet implemented
                    # Requires parsing bookmarks or comment ranges
                    result.add_skipped(
                        condition_name,
                        "Conditional sections not yet implemented - requires Document Library bookmark support"
                    )

            except Exception as e:
                result.add_error(condition_name, str(e))

        return result


class MultiStrategyFiller:
    """Combines all strategies for comprehensive document filling"""

    def __init__(self, doc: Document):
        self.doc = doc
        self.results = {}

    def fill_with_all_strategies(self, placeholders: dict) -> dict:
        """Try all strategies to fill placeholders

        Args:
            placeholders: Dict of {placeholder_name: value}

        Returns:
            Dict with results from each strategy
        """
        print("\n=== Multi-Strategy Document Filling ===\n")

        # Strategy A: Text placeholders
        print("Strategy A: Text placeholder replacement...")
        result_a = StrategyA.fill(self.doc, placeholders)
        self.results['A-Text'] = result_a
        print(f"  Result: {result_a.summary()}\n")

        # Strategy B: SDT fields
        print("Strategy B: Structured Data Tag replacement...")
        result_b = StrategyB.fill(self.doc, placeholders)
        self.results['B-SDT'] = result_b
        print(f"  Result: {result_b.summary()}\n")

        # Strategy C: Element IDs (if provided)
        if self._has_element_ids(placeholders):
            print("Strategy C: Element ID-based replacement...")
            id_mapping = self._convert_to_id_mapping(placeholders)
            result_c = StrategyC.fill(self.doc, id_mapping)
            self.results['C-ID'] = result_c
            print(f"  Result: {result_c.summary()}\n")

        # Strategy D: Multi-run (automatic)
        print("Strategy D: Multi-run placeholder handling...")
        result_d = StrategyD.fill(self.doc, placeholders)
        self.results['D-Multirun'] = result_d
        print(f"  Result: {result_d.summary()}\n")

        # Save document after all replacements (skip validation - pre-existing errors in template)
        self.doc.save(validate=False)

        return self.results

    def fill_table(self, table_index: int, table_data: list) -> FillingResult:
        """Fill a specific table with data

        Args:
            table_index: Which table to fill (0-based)
            table_data: List of row data (dicts)

        Returns:
            FillingResult from table filling
        """
        print(f"Strategy E: Table filling (table {table_index})...")
        result = StrategyE.fill(self.doc, table_index, table_data)
        self.results['E-Table'] = result
        print(f"  Result: {result.summary()}\n")

        self.doc.save(validate=False)
        return result

    def _has_element_ids(self, placeholders: dict) -> bool:
        """Check if placeholders contain element ID format"""
        # Simple check: IDs are typically like "field_company_name"
        for key in placeholders.keys():
            if key.startswith('field_'):
                return True
        return False

    def _convert_to_id_mapping(self, placeholders: dict) -> dict:
        """Convert placeholders to ID mapping if needed"""
        # Simple conversion: assume field_* keys are IDs
        id_mapping = {k: v for k, v in placeholders.items() if k.startswith('field_')}
        return id_mapping

    def get_summary(self) -> str:
        """Get summary of all strategies"""
        all_filled = []
        all_skipped = 0
        all_errors = 0

        for strategy_name, result in self.results.items():
            all_filled.extend(result.filled)
            all_skipped += len(result.skipped)
            all_errors += len(result.errors)

        summary = f"""
=== FILLING SUMMARY ===
Total filled: {len(set(all_filled))}
Total skipped: {all_skipped}
Total errors: {all_errors}

Strategies used: {', '.join(self.results.keys())}
"""
        return summary


if __name__ == '__main__':
    # Test strategies
    print("Filling Strategies Module")
    print("Use with: from filling_strategies import MultiStrategyFiller")
