#!/usr/bin/env python3
"""Unpack DOCX to XML using local ooxml_scripts"""
import sys
import os
import subprocess

current_dir = os.path.dirname(os.path.abspath(__file__))
unpack_script = os.path.join(current_dir, '..', 'ooxml_scripts', 'unpack.py')

if not os.path.exists(unpack_script):
    print(f"ERROR: unpack.py not found at {unpack_script}", file=sys.stderr)
    sys.exit(1)

try:
    result = subprocess.run(
        [sys.executable, unpack_script] + sys.argv[1:],
        check=True
    )
    sys.exit(result.returncode)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
