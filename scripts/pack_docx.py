#!/usr/bin/env python3
"""Pack unpacked XML directory back to DOCX using local ooxml_scripts"""
import sys
import os
import subprocess

current_dir = os.path.dirname(os.path.abspath(__file__))
pack_script = os.path.join(current_dir, '..', 'ooxml_scripts', 'pack.py')

if not os.path.exists(pack_script):
    print(f"ERROR: pack.py not found at {pack_script}", file=sys.stderr)
    sys.exit(1)

try:
    result = subprocess.run(
        [sys.executable, pack_script] + sys.argv[1:],
        check=True
    )
    sys.exit(result.returncode)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
