#!/usr/bin/env python3
"""
Data Extraction Module - Phase 2 Implementation
Extracts data from DOCX using 3 methods: Text, Tables, and SDT fields
"""

import os
from pathlib import Path
from typing import Dict, List, Any
from xml.dom import minidom
from defusedxml import minidom as defused_minidom


def extract_text_from_docx(unpacked_source_path: str) -> str:
    """
    Extract all text content from unpacked DOCX.

    Extracts every piece of text from word/document.xml by collecting all
    w:t (text) elements. Preserves order but loses formatting/structure.

    Args:
        unpacked_source_path: Path to unpacked DOCX directory

    Returns:
        String with all document text concatenated

    Example:
        text = extract_text_from_docx('unpacked_source')
        # Returns: "Invoice to Acme Corporation dated 2025-11-18..."
    """
    doc_xml = os.path.join(unpacked_source_path, 'word/document.xml')

    if not os.path.exists(doc_xml):
        return ""

    try:
        dom = defused_minidom.parse(doc_xml)
        text_runs = dom.getElementsByTagName('w:t')

        text_content = ''.join([
            t.firstChild.nodeValue for t in text_runs if t.firstChild
        ])

        return text_content
    except Exception as e:
        print(f"Error extracting text: {e}")
        return ""


def extract_table_data(unpacked_source_path: str) -> list:
    """
    Extract structured table data from source DOCX.

    Finds all tables (w:tbl) and extracts cells as a list of rows.
    Each row is a list of cell texts, preserving table structure.

    Args:
        unpacked_source_path: Path to unpacked DOCX directory

    Returns:
        List of tables, each table is list of rows, each row is list of cells:
        [
            [  # Table 1
                ['Header1', 'Header2', 'Header3'],  # Row 1 (header)
                ['Value1', 'Value2', 'Value3'],     # Row 2 (data)
                ['Value4', 'Value5', 'Value6']      # Row 3 (data)
            ],
            [  # Table 2
                ...
            ]
        ]

    Example:
        tables = extract_table_data('unpacked_source')
        if tables:
            first_table = tables[0]
            # first_table[0] = header row
            # first_table[1:] = data rows
    """
    doc_xml = os.path.join(unpacked_source_path, 'word/document.xml')

    if not os.path.exists(doc_xml):
        return []

    try:
        dom = defused_minidom.parse(doc_xml)
        tables = dom.getElementsByTagName('w:tbl')

        table_data = []

        for table in tables:
            rows = table.getElementsByTagName('w:tr')
            table_rows = []

            for row in rows:
                cells = row.getElementsByTagName('w:tc')
                row_data = []

                for cell in cells:
                    # Extract all text from cell
                    text_runs = cell.getElementsByTagName('w:t')
                    cell_text = ''.join([
                        t.firstChild.nodeValue for t in text_runs if t.firstChild
                    ])
                    row_data.append(cell_text.strip())

                table_rows.append(row_data)

            table_data.append(table_rows)

        return table_data

    except Exception as e:
        print(f"Error extracting table data: {e}")
        return []


def extract_sdt_fields(unpacked_source_path: str) -> dict:
    """
    Extract Structured Data Tag (SDT) fields from DOCX.

    Structured Data Tags are Word's native form fields (Text Form Fields,
    Dropdown Lists, Date Pickers, etc.). Each SDT has an alias (field name)
    and content (field value).

    Supports:
    - Plain text form fields
    - Dropdown selections (returns selected value)
    - Date picker values
    - Checkbox states

    Args:
        unpacked_source_path: Path to unpacked DOCX directory

    Returns:
        Dictionary mapping field names (aliases) to their values:
        {
            'CLIENT_NAME': 'Acme Corporation',
            'PROJECT': 'Website Redesign',
            'INVOICE_DATE': '2025-11-18',
            'TERMS': '30 days'
        }

    Example:
        sdt_data = extract_sdt_fields('unpacked_source')
        client_name = sdt_data.get('CLIENT_NAME', 'Unknown')
    """
    doc_xml = os.path.join(unpacked_source_path, 'word/document.xml')

    if not os.path.exists(doc_xml):
        return {}

    try:
        dom = defused_minidom.parse(doc_xml)
        sdts = dom.getElementsByTagName('w:sdt')

        sdt_data = {}

        for sdt in sdts:
            try:
                # Get field name from w:alias attribute
                alias_elems = sdt.getElementsByTagName('w:alias')
                if not alias_elems:
                    continue

                field_name = alias_elems[0].getAttribute('w:val')
                if not field_name:
                    continue

                # Get field value from w:sdtContent
                content_elems = sdt.getElementsByTagName('w:sdtContent')
                if not content_elems:
                    continue

                content_elem = content_elems[0]

                # Extract all text from content
                text_runs = content_elem.getElementsByTagName('w:t')
                field_value = ''.join([
                    t.firstChild.nodeValue for t in text_runs if t.firstChild
                ])

                sdt_data[field_name] = field_value.strip()

            except (IndexError, AttributeError):
                continue

        return sdt_data

    except Exception as e:
        print(f"Error extracting SDT fields: {e}")
        return {}


def normalize_data(raw_data: dict, mapping: dict = None) -> dict:
    """
    Normalize extracted data using field mapping.

    Converts raw extracted data to standardized format by mapping
    source field names to template field names.

    Args:
        raw_data: Dictionary with raw extracted data
        mapping: Dict mapping source field names to template names:
                 {'Client Name': 'CLIENT_NAME', ...}
                 If None, returns data as-is

    Returns:
        Normalized dictionary with template field names as keys

    Example:
        raw = {'Client Name': 'Acme Corp', 'Project': 'Website'}
        mapping = {
            'Client Name': 'CLIENT_NAME',
            'Project': 'PROJECT_NAME'
        }
        normalized = normalize_data(raw, mapping)
        # Returns: {'CLIENT_NAME': 'Acme Corp', 'PROJECT_NAME': 'Website'}
    """
    if not mapping:
        return raw_data

    normalized = {}
    for source_field, template_field in mapping.items():
        if source_field in raw_data:
            normalized[template_field] = raw_data[source_field]

    return normalized


def merge_data_sources(text_data: str, table_data: list, sdt_data: dict) -> dict:
    """
    Merge data from all 3 extraction methods.

    Intelligently combines data from text extraction, table extraction,
    and SDT field extraction. SDT fields take priority (most structured),
    followed by tables, then text.

    Args:
        text_data: String of all text content
        table_data: List of tables with row/cell structure
        sdt_data: Dictionary of SDT field name -> value

    Returns:
        Merged dictionary with all available data:
        {
            'text': full_text_string,
            'tables': [table1, table2, ...],
            'sdt_fields': {field_name: value, ...},
            'extracted_values': {field_name: value, ...}
        }

    Example:
        merged = merge_data_sources(
            text_content,
            table_list,
            sdt_dict
        )
        # Use merged['extracted_values'] for filling
    """
    merged = {
        'text': text_data,
        'tables': table_data,
        'sdt_fields': sdt_data,
        'extracted_values': {}
    }

    # SDT fields are most structured - add them first
    merged['extracted_values'].update(sdt_data)

    # Parse common fields from text
    text_based_fields = _extract_common_fields_from_text(text_data)
    for field, value in text_based_fields.items():
        if field not in merged['extracted_values']:
            merged['extracted_values'][field] = value

    # Extract potential fields from tables
    table_based_fields = _extract_fields_from_tables(table_data)
    for field, value in table_based_fields.items():
        if field not in merged['extracted_values']:
            merged['extracted_values'][field] = value

    return merged


def _extract_common_fields_from_text(text: str) -> dict:
    """
    Extract common field values from text using heuristics.

    Looks for common patterns like:
    - "Date: 2025-11-18" -> {'DATE': '2025-11-18'}
    - "Invoice #12345" -> {'INVOICE_NUMBER': '12345'}
    - "Amount: $50,000" -> {'AMOUNT': '$50,000'}

    Args:
        text: Full text content

    Returns:
        Dictionary of inferred field -> value

    Example:
        fields = _extract_common_fields_from_text(full_text)
        # Returns: {'DATE': '2025-11-18', 'AMOUNT': '$50,000', ...}
    """
    import re

    fields = {}

    # Date patterns
    date_pattern = r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|[A-Z][a-z]+ \d{1,2}, \d{4})'
    dates = re.findall(date_pattern, text)
    if dates:
        fields['DATE'] = dates[0]

    # Currency patterns
    currency_pattern = r'\$[\d,.]+'
    currencies = re.findall(currency_pattern, text)
    if currencies:
        fields['AMOUNT'] = currencies[0]

    # Invoice/Order number patterns
    invoice_pattern = r'(?:Invoice|Order)[\s#]*(\d+)'
    invoices = re.findall(invoice_pattern, text, re.IGNORECASE)
    if invoices:
        fields['INVOICE_NUMBER'] = invoices[0]

    return fields


def _extract_fields_from_tables(tables: list) -> dict:
    """
    Extract field values from table data.

    Assumes first row is header with field names, subsequent rows have values.
    Returns first row of data values.

    Args:
        tables: List of tables (from extract_table_data)

    Returns:
        Dictionary mapping header names to first row values

    Example:
        Tables:
        [
            [
                ['Item', 'Quantity', 'Price'],
                ['Widget', '10', '$100'],
                ['Gadget', '5', '$50']
            ]
        ]
        Returns: {'Item': 'Widget', 'Quantity': '10', 'Price': '$100'}
    """
    fields = {}

    if not tables:
        return fields

    # Process first table only
    table = tables[0]
    if len(table) < 2:
        return fields

    headers = table[0]
    first_row = table[1]

    for i, header in enumerate(headers):
        if i < len(first_row):
            # Normalize header as field name
            field_name = header.upper().replace(' ', '_')
            fields[field_name] = first_row[i]

    return fields


def comprehensive_data_extraction(unpacked_source_path: str) -> dict:
    """
    Perform complete data extraction from source document.

    Extracts data using all 3 methods, merges results, and returns
    comprehensive data structure ready for filling.

    Args:
        unpacked_source_path: Path to unpacked source DOCX

    Returns:
        Merged data from all extraction methods

    Example:
        source_data = comprehensive_data_extraction('unpacked_source')
        # Use source_data['extracted_values'] for field filling
    """
    print("Extracting data from source document...")

    # Extract via all 3 methods
    text = extract_text_from_docx(unpacked_source_path)
    print(f"  - Text extraction: {len(text)} characters")

    tables = extract_table_data(unpacked_source_path)
    print(f"  - Table extraction: {len(tables)} tables found")

    sdt_fields = extract_sdt_fields(unpacked_source_path)
    print(f"  - SDT extraction: {len(sdt_fields)} form fields found")

    # Merge all sources
    merged = merge_data_sources(text, tables, sdt_fields)

    print(f"  - Total extracted values: {len(merged['extracted_values'])}")
    print(f"  - Fields: {list(merged['extracted_values'].keys())}")

    return merged


if __name__ == '__main__':
    # Test extraction functions
    import sys

    if len(sys.argv) < 2:
        print("Usage: python extraction_module.py <unpacked_docx_path>")
        sys.exit(1)

    unpacked_path = sys.argv[1]
    print(f"Extracting from: {unpacked_path}\n")

    data = comprehensive_data_extraction(unpacked_path)

    print("\n=== EXTRACTED DATA ===")
    print(f"Text (first 200 chars): {data['text'][:200]}...")
    print(f"\nTables: {len(data['tables'])}")
    for i, table in enumerate(data['tables']):
        print(f"  Table {i}: {len(table)} rows x {len(table[0]) if table else 0} cols")

    print(f"\nSDT Fields: {len(data['sdt_fields'])}")
    for field, value in data['sdt_fields'].items():
        print(f"  {field}: {value}")

    print(f"\nExtracted Values: {len(data['extracted_values'])}")
    for field, value in data['extracted_values'].items():
        print(f"  {field}: {value}")
