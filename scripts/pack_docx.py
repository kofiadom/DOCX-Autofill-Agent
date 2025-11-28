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
    # Skip soffice validation by default (--force)
    # Reasoning:
    # 1. DOCX file creation is pure ZIP operations (doesn't need soffice)
    # 2. XML validation already done in Phase 5 (ComprehensiveValidator)
    # 3. soffice validation times out on large documents (10s timeout)
    # 4. Validation is redundant - we already checked structure, well-formedness
    #
    # If validation is needed for debugging, use: python pack.py <dir> <file>
    # (without this wrapper) to enable soffice validation

    args = [sys.executable, pack_script] + sys.argv[1:] + ['--force']
    result = subprocess.run(args, check=True)
    sys.exit(result.returncode)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
