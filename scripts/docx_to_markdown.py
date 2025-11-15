#!/usr/bin/env python3
"""Convert DOCX to markdown for agent to analyze"""
import sys
import subprocess
import os

def docx_to_markdown(docx_path, output_md_path):
    """
    Convert DOCX to markdown using pandoc.

    Args:
        docx_path: Path to .docx file
        output_md_path: Path where to save .md file

    Returns:
        0 on success, 1 on failure
    """
    # Check if pandoc is installed
    try:
        result = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
    except FileNotFoundError:
        print("ERROR: pandoc not found. Install pandoc", file=sys.stderr)
        return 1

    # Convert DOCX to markdown
    try:
        result = subprocess.run(
            ["pandoc", docx_path, "-o", output_md_path],
            capture_output=True,
            text=True,
            check=True
        )
        print(f"Successfully converted {docx_path} to markdown")
        return 0
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Conversion failed: {e.stderr}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        return 1

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python docx_to_markdown.py <docx_file> <output_md>")
        sys.exit(1)

    sys.exit(docx_to_markdown(sys.argv[1], sys.argv[2]))
