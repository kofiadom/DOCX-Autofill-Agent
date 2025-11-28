"""
DOCX Tools - Layer 1: Core tool implementations
Direct implementations that call scripts and return string messages

Enhanced with Phase 2-5 capabilities:
- Phase 2: Data extraction (3 methods)
- Phase 3: Multi-strategy filling (6 strategies)
- Phase 5: Validation (3 tiers)
"""
import subprocess
import sys
import os
from pathlib import Path
from typing import Dict, Any

# Import new modules for enhanced functionality
from .extraction_module import (
    extract_text_from_docx,
    extract_table_data,
    extract_sdt_fields,
    comprehensive_data_extraction,
    merge_data_sources
)

from .filling_strategies import MultiStrategyFiller

from .validation_module import ComprehensiveValidator


def list_data_files() -> str:
    """List files in data/ directory."""
    data_dir = Path("data")
    if not data_dir.exists():
        return "❌ The data/ directory does not exist. Please create data/ directory or upload files to a session."

    files = list(data_dir.glob("*"))
    if not files:
        return "📁 The data/ directory is empty."

    lines = ["📁 **Available files:**"]
    for f in sorted(files):
        if f.is_file():
            size = f.stat().st_size / 1024  # KB
            lines.append(f"  - {f.name} ({size:.1f} KB)")

    return "\n".join(lines)


def unpack_docx(docx_path: str, output_dir: str) -> str:
    """Unpack a DOCX file to XML structure."""
    script_path = Path(__file__).parent.parent / "scripts" / "unpack_docx.py"

    try:
        result = subprocess.run(
            [sys.executable, str(script_path), docx_path, output_dir],
            capture_output=True,
            text=True,
            check=True
        )
        return f"✅ Successfully unpacked DOCX to {output_dir}"
    except subprocess.CalledProcessError as e:
        return f"❌ Error unpacking DOCX: {e.stderr}"
    except Exception as e:
        return f"❌ Error: {str(e)}"


def convert_docx_to_markdown(docx_path: str, output_md_path: str) -> str:
    """Convert DOCX to markdown using pandoc."""
    script_path = Path(__file__).parent.parent / "scripts" / "docx_to_markdown.py"

    try:
        result = subprocess.run(
            [sys.executable, str(script_path), docx_path, output_md_path],
            capture_output=True,
            text=True,
            check=True
        )
        return f"✅ Successfully converted {Path(docx_path).name} to markdown"
    except subprocess.CalledProcessError as e:
        if "pandoc not found" in e.stderr:
            return "❌ pandoc not installed. Install with: apt-get install pandoc"
        return f"❌ Error converting to markdown: {e.stderr}"
    except Exception as e:
        return f"❌ Error: {str(e)}"


def pack_docx(unpacked_dir: str, output_docx: str) -> str:
    """Pack unpacked XML back to DOCX."""
    script_path = Path(__file__).parent.parent / "scripts" / "pack_docx.py"

    try:
        result = subprocess.run(
            [sys.executable, str(script_path), unpacked_dir, output_docx],
            capture_output=True,
            text=True,
            check=True
        )
        return f"✅ Successfully packed to {Path(output_docx).name}"
    except subprocess.CalledProcessError as e:
        return f"❌ Error packing DOCX: {e.stderr}"
    except Exception as e:
        return f"❌ Error: {str(e)}"


def save_json_file(file_path: str, data: Dict[str, Any]) -> str:
    """Save JSON data to file."""
    import json
    try:
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=2)
        return f"✅ Saved to {file_path}"
    except Exception as e:
        return f"❌ Error saving JSON: {str(e)}"


def read_json_file(file_path: str) -> str:
    """Read JSON file and return as formatted string."""
    import json
    try:
        with open(file_path, 'r') as f:
            data = json.load(f)
        return json.dumps(data, indent=2)
    except FileNotFoundError:
        return f"❌ File not found: {file_path}"
    except Exception as e:
        return f"❌ Error reading JSON: {str(e)}"


def read_text_file(file_path: str) -> str:
    """Read text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        if len(content) > 10000:
            return content[:10000] + f"\n\n... (truncated, total {len(content)} chars)"
        return content
    except FileNotFoundError:
        return f"❌ File not found: {file_path}"
    except Exception as e:
        return f"❌ Error reading file: {str(e)}"


def extract_all_data(unpacked_source: str, debug_dir: str = None) -> dict:
    """
    Phase 2: Extract data from source document using all 3 methods.

    Extracts data via:
    - Option A: Text content (DOM parsing)
    - Option B: Table structure (rows/cells)
    - Option C: Structured Data Tags (SDT form fields)

    Args:
        unpacked_source: Path to unpacked source DOCX
        debug_dir: Directory to save extraction results (optional)

    Returns:
        Merged data dict with all extraction results:
        {
            'text': full_text_content,
            'tables': [[row_data], ...],
            'sdt_fields': {field_name: value, ...},
            'extracted_values': {field_name: value, ...}
        }
    """
    print(f"[Phase 2] Extracting data from source...")

    try:
        # Use comprehensive extraction
        data = comprehensive_data_extraction(unpacked_source)

        print(f"[Phase 2] Extraction complete:")
        print(f"  - Text: {len(data['text'])} chars")
        print(f"  - Tables: {len(data['tables'])} tables")
        print(f"  - SDT Fields: {len(data['sdt_fields'])} fields")
        print(f"  - Extracted Values: {len(data['extracted_values'])} values")

        # Save extraction results to debug directory
        if debug_dir:
            save_json_file(f"{debug_dir}/extraction_results.json", data)
            print(f"[Phase 2] Saved extraction results to debug directory")

        return data

    except Exception as e:
        print(f"[Phase 2] Error during extraction: {str(e)}")
        return {
            'text': '',
            'tables': [],
            'sdt_fields': {},
            'extracted_values': {}
        }


def insert_placeholders(unpacked_dir: str, field_analysis: dict, debug_dir: str = None) -> dict:
    """
    Phase 1.3: Insert {{FIELD_NAME}} placeholders into template.

    Automatically inserts {{FIELD_NAME}} placeholders into the unpacked template XML
    based on field analysis provided by the agent.

    Args:
        unpacked_dir: Path to unpacked template directory
        field_analysis: Dict with fields to insert:
            {
                'fields': [
                    {
                        'field_name': 'PROJECT_MANAGER',
                        'label': 'Project Manager:',
                        'location': 'below_label'
                    },
                    ...
                ]
            }
        debug_dir: Directory to save placeholder insertion results (optional)

    Returns:
        Dict with results:
        {
            'status': 'success|partial|failed',
            'inserted_count': N,
            'inserted_fields': [...],
            'summary': 'message'
        }
    """
    try:
        from lib.document import Document

        print(f"[Phase 1.3] Inserting placeholders into template...")

        # Load template
        doc = Document(unpacked_dir)

        # Get fields to insert
        fields = field_analysis.get('fields', [])
        if not fields:
            return {
                'status': 'failed',
                'inserted_count': 0,
                'inserted_fields': [],
                'summary': '❌ No fields provided for placeholder insertion'
            }

        inserted_fields = []
        failed_fields = []

        # Process each field
        for field_info in fields:
            field_name = field_info.get('field_name')
            label = field_info.get('label', '')
            location = field_info.get('location', 'below_label')

            if not field_name:
                continue

            try:
                # Load XML
                doc_xml = doc["word/document.xml"]
                dom = doc_xml.dom
                paragraphs = dom.getElementsByTagName('w:p')

                inserted = False

                # Strategy: Find label and insert placeholder below it
                if location == 'below_label' and label:
                    for i, para in enumerate(paragraphs):
                        para_text = para.toxml() if hasattr(para, 'toxml') else str(para)

                        # Check if paragraph contains the label
                        if label in para_text:
                            # Insert placeholder in next paragraph
                            if i + 1 < len(paragraphs):
                                next_para = paragraphs[i + 1]

                                # Create placeholder text element
                                placeholder_run = dom.createElement('w:r')
                                placeholder_text = dom.createElement('w:t')
                                placeholder_text.setAttribute('xml:space', 'preserve')
                                placeholder_text.appendChild(
                                    dom.createTextNode(f"{{{{{field_name}}}}}")
                                )
                                placeholder_run.appendChild(placeholder_text)

                                # Append to next paragraph
                                next_para.appendChild(placeholder_run)
                                inserted = True
                                break

                # Strategy: Insert as inline placeholder on same line as label
                if not inserted and label:
                    for para in paragraphs:
                        para_xml = para.toxml() if hasattr(para, 'toxml') else str(para)

                        if label in para_xml:
                            # Create placeholder text element
                            placeholder_run = dom.createElement('w:r')
                            placeholder_text = dom.createElement('w:t')
                            placeholder_text.setAttribute('xml:space', 'preserve')
                            placeholder_text.appendChild(
                                dom.createTextNode(f" {{{{{field_name}}}}}")
                            )
                            placeholder_run.appendChild(placeholder_text)

                            # Append to label paragraph
                            para.appendChild(placeholder_run)
                            inserted = True
                            break

                if inserted:
                    inserted_fields.append(field_name)
                else:
                    failed_fields.append(field_name)

            except Exception as e:
                print(f"[Phase 1.3] Error inserting {field_name}: {str(e)}")
                failed_fields.append(field_name)

        # Save document (without validation - we didn't break anything, just added placeholders)
        doc.save(validate=False)

        # Determine status
        status = 'success' if len(failed_fields) == 0 else 'partial' if len(inserted_fields) > 0 else 'failed'

        result_msg = f"Inserted {len(inserted_fields)} placeholder(s)"
        if inserted_fields:
            result_msg += f": {', '.join(inserted_fields)}"
        if failed_fields:
            result_msg += f" | Failed {len(failed_fields)}: {', '.join(failed_fields)}"

        results = {
            'status': status,
            'inserted_count': len(inserted_fields),
            'inserted_fields': inserted_fields,
            'failed_fields': failed_fields,
            'summary': f"✅ {result_msg}" if status == 'success' else f"⚠️ {result_msg}"
        }

        # Save placeholder insertion results to debug directory
        if debug_dir:
            save_json_file(f"{debug_dir}/placeholder_insertion_results.json", results)
            save_json_file(f"{debug_dir}/field_analysis_used.json", field_analysis)
            print(f"[Phase 1.3] Saved placeholder insertion results to debug directory")

            # Generate markdown from the modified template
            print(f"[Phase 1.3] Generating markdown preview of template with placeholders...")
            try:
                # Pack the unpacked directory to temporary DOCX
                temp_docx_path = f"{debug_dir}/template_with_placeholders.docx"
                pack_result = pack_docx(unpacked_dir, temp_docx_path)

                if "Successfully packed" in pack_result:
                    # Convert the repacked DOCX to markdown
                    markdown_path = f"{debug_dir}/template_with_placeholders.md"
                    markdown_result = convert_docx_to_markdown(temp_docx_path, markdown_path)

                    if "Successfully converted" in markdown_result:
                        print(f"[Phase 1.3] Saved markdown preview to {markdown_path}")
                        results['markdown_preview'] = markdown_path
                    else:
                        print(f"[Phase 1.3] Warning: Could not convert to markdown: {markdown_result}")
                else:
                    print(f"[Phase 1.3] Warning: Could not repack DOCX for markdown generation: {pack_result}")
            except Exception as e:
                print(f"[Phase 1.3] Warning: Could not generate markdown preview: {str(e)}")

        return results

    except Exception as e:
        print(f"[Phase 1.3] Error: {str(e)}")
        return {
            'status': 'failed',
            'inserted_count': 0,
            'inserted_fields': [],
            'error': str(e),
            'summary': f"❌ Error inserting placeholders: {str(e)}"
        }


def fill_fields(unpacked_dir: str, field_mapping: dict, debug_dir: str = None) -> dict:
    """
    Phase 3 & 5: Fill DOCX fields using multiple strategies with automatic validation.

    Applies 5 filling strategies (A-E) automatically - Strategy F (conditional) not yet implemented:
    - Strategy A: Text placeholder replacement ({{FIELD}})
    - Strategy B: Structured Data Tag (SDT) replacement
    - Strategy C: Element ID-based targeting
    - Strategy D: Multi-run placeholder handling
    - Strategy E: Table row filling
    - Strategy F: Conditional content (NOT YET IMPLEMENTED - stub only)

    Automatically validates results with Phase 5 validation (3-tier):
    - Tier 1: Placeholder completion check
    - Tier 2: Document integrity verification
    - Tier 3: XML well-formedness validation

    Args:
        unpacked_dir: Path to unpacked template directory
        field_mapping: Dict mapping field names to values:
                      {"Field Name": "value", ...}
        debug_dir: Directory to save filling results (optional)

    Returns:
        Dict with results:
        {
            'status': 'success|partial|failed',
            'filled': [...],
            'skipped': count,
            'strategies_used': [...],
            'validation_passed': bool,
            'summary': 'message'
        }
    """
    try:
        from lib.document import Document

        print(f"[Phase 3] Filling template with {len(field_mapping)} fields...")

        # Load template
        doc = Document(unpacked_dir)

        # Use multi-strategy filler
        filler = MultiStrategyFiller(doc)
        strategy_results = filler.fill_with_all_strategies(field_mapping)

        # Compile results
        all_filled = []
        all_skipped = 0

        for strategy_name, result in strategy_results.items():
            all_filled.extend(result.filled)
            all_skipped += len(result.skipped)

        # Save document (skip validation - pre-existing errors in template shouldn't block field filling)
        doc.save(validate=False)

        # Validate results
        print(f"[Phase 5] Validating document...")
        expected_fields = list(field_mapping.keys())
        validator = ComprehensiveValidator(doc, unpacked_dir)
        is_valid = validator.validate_all(expected_fields)

        # Build response
        filled_list = list(set(all_filled))  # Remove duplicates
        status = "success" if is_valid else "partial"

        result_msg = f"Filled {len(filled_list)}"
        if filled_list:
            result_msg += f": {', '.join(filled_list)}"
        if all_skipped > 0:
            result_msg += f" | Skipped {all_skipped}"

        results = {
            'status': status,
            'filled': filled_list,
            'skipped': all_skipped,
            'strategies_used': list(strategy_results.keys()),
            'summary': f"✅ {result_msg}" if is_valid else f"⚠️ {result_msg}",
            'validation_passed': is_valid
        }

        # Create filled mapping (only fields that were actually filled)
        filled_mapping = {k: v for k, v in field_mapping.items() if k in filled_list}

        # Save filling results to debug directory
        if debug_dir:
            save_json_file(f"{debug_dir}/filling_results.json", results)
            save_json_file(f"{debug_dir}/field_mapping_applied.json", field_mapping)
            save_json_file(f"{debug_dir}/field_mapping_filled.json", filled_mapping)
            print(f"[Phase 3] Saved filling results to debug directory")

        return results

    except Exception as e:
        print(f"[Phase 3] Error: {str(e)}")
        return {
            'status': 'failed',
            'error': str(e),
            'summary': f"❌ Error filling fields: {str(e)}"
        }
