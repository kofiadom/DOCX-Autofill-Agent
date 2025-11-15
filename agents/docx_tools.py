"""
DOCX Tools - Layer 1: Core tool implementations
Direct implementations that call scripts and return string messages
"""
import subprocess
import sys
import os
from pathlib import Path
from typing import Dict, Any


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


def fill_fields(unpacked_dir: str, field_mapping: dict) -> str:
    """
    Fill DOCX fields using Document Library and label-based field detection.

    Algorithm:
    1. For each field in mapping:
       - Find paragraph containing label text
       - Search forward for empty field (empty <w:t> node)
       - Fill with value
    2. Save document

    Args:
        unpacked_dir: Path to unpacked DOCX directory
        field_mapping: Dict of {"label": "value"}

    Returns:
        Status message with fill results
    """
    try:
        from lib.document import Document

        doc = Document(unpacked_dir)
        x = doc["word/document.xml"]

        filled = []
        skipped = []

        for label, value in field_mapping.items():
            try:
                # Find label paragraph
                label_para = x.get_node(tag="w:p", contains=label)

                # Get all paragraphs
                all_paras = x.dom.getElementsByTagName("w:p")
                label_idx = None

                # Find index of label paragraph
                for i, para in enumerate(all_paras):
                    if para == label_para:
                        label_idx = i
                        break

                if label_idx is None:
                    skipped.append(label)
                    continue

                # Search forward for empty field
                empty_found = False
                for search_para in all_paras[label_idx:]:
                    t_nodes = search_para.getElementsByTagName("w:t")
                    for t_node in t_nodes:
                        text = t_node.firstChild.data if t_node.firstChild else ""
                        if text.strip() == "":
                            t_node.firstChild.data = value
                            filled.append(label)
                            empty_found = True
                            break
                    if empty_found:
                        break

                if not empty_found:
                    skipped.append(label)

            except ValueError:
                skipped.append(label)

        # Save document
        doc.save()

        # Build result message
        result = "Filled {} field(s)".format(len(filled))
        if filled:
            result += ": " + ", ".join(filled)
        if skipped:
            result += " | Skipped {}: {}".format(len(skipped), ", ".join(skipped))

        return "✅ " + result

    except Exception as e:
        return "❌ Error filling fields: {}".format(str(e))
