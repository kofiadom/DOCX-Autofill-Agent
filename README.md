# DOCX Autofill Agent with Agno

A production-ready FastAPI web service that intelligently fills DOCX templates with data extracted from source documents using Claude AI.

**Status:** ✅ Ready to Launch
**Version:** 1.0
**Python:** 3.8+

## Workflow Overview

The autofill process follows a **three-stage workflow** with clear separation between automation and LLM analysis:

```
┌──────────────────────┐
│  STAGE 1: AUTOMATION │
│      (Unpack)        │
└──────┬───────────────┘
       │ unpack_docx_script(template.docx)
       ↓
    workspaces/{session}/unpacked/template/
    ├── word/document.xml
    ├── word/styles.xml
    └── [all DOCX XML files preserved]

┌──────────────────────────┐
│  STAGE 2: LLM ANALYSIS   │
│  (Analyze & Map)         │
└──────┬───────────────────┘
       │ Agent reads: word/document.xml
       │ Agent finds: {{placeholder_name}} patterns
       │ Agent creates: replacements.json mapping
       ↓
    workspaces/{session}/debug/replacements.json
    {
      "employee_name": "John Smith",
      "employee_id": "EMP12345"
    }

┌──────────────────────┐
│  STAGE 3: AUTOMATION │
│  (Fill & Pack)       │
└──────┬───────────────┘
       │ fill_docx_script(unpacked/, replacements.json)
       │ Replaces {{placeholder}} with values
       │ Preserves all formatting, styles, structure
       ↓
       │ pack_docx_script(unpacked/)
       │ Converts back to DOCX ZIP format
       ↓
    workspaces/{session}/output/filled.docx
```

## Key Features

✅ **Standalone**: Includes all necessary OOXML utilities (unpack/pack/document library)
✅ **Formatting Preserved**: Document structure, styles, fonts, colors maintained
✅ **Session Isolated**: Each user gets private workspace (input/unpacked/debug/output)
✅ **Two-Way Integration**: Can extract data from source documents OR accept manual mappings
✅ **XML-Level Precision**: Uses Document Library for safe, precise text replacements
✅ **No External Dependencies**: Only requires `defusedxml` (minimal, security-focused)

## Directory Structure

```
Agno-DOCX-Autofill-Agent/
├── agents/
│   ├── __init__.py
│   ├── session_workspace.py      # Workspace manager for session isolation
│   └── session_tools.py          # Tool implementations (unpack, find, fill, pack)
│
├── scripts/                      # Automation scripts called by tools
│   ├── unpack_docx.py            # Unpack DOCX to XML
│   ├── extract_placeholders.py   # Find {{placeholders}}
│   ├── fill_docx.py              # Replace {{placeholders}} with values
│   └── pack_docx.py              # Pack XML back to DOCX
│
├── ooxml_scripts/                # Local copy of document-skills OOXML utilities
│   ├── unpack.py                 # Core unpacking logic
│   ├── pack.py                   # Core packing logic
│   ├── validate.py               # Validation utilities
│   └── validation/               # Schema validators
│
├── lib/                          # Document Library from document-skills
│   ├── utilities.py              # XMLEditor class for XML manipulation
│   └── document.py               # High-level Document API
│
├── templates/                    # Template XML files for advanced features
│   ├── comments.xml
│   ├── commentsExtended.xml
│   └── [other template files]
│
├── workspaces/                   # Session directories (auto-created)
│   ├── {session_id}/
│   │   ├── input/                # User uploaded files
│   │   ├── unpacked/             # Unpacked DOCX work directory
│   │   ├── debug/                # Analysis files (replacements.json)
│   │   └── output/               # Final filled documents
│
├── requirements.txt
└── README.md
```

## Installation & Setup

### 1. Install Dependencies

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Basic Usage

```python
from agents import SessionTools

# Create tools for a session
tools = SessionTools(session_id="user_123")

# STAGE 1: AUTOMATION - Unpack template
result = tools.unpack_template("template.docx")
# Output: unpacked/template/ directory with all XML files

# STAGE 2: LLM ANALYSIS - Find placeholders
result = tools.find_placeholders()
# Output: {"placeholders": ["employee_name", "employee_id", ...]}

# Agent creates mapping from source data or user input
tools.create_replacement_mapping({
    "employee_name": "John Smith",
    "employee_id": "EMP12345"
})

# STAGE 3: AUTOMATION - Fill and pack
tools.fill_template()  # Replaces {{placeholder}} with values
tools.pack_template()  # Converts back to DOCX
# Output: output/template_filled.docx
```

## The Three Stages Explained

### Stage 1: AUTOMATION - Unpack

**What happens:**
- DOCX file (ZIP archive) is extracted
- All XML files are pretty-printed for readability
- Document structure preserved exactly

**Scripts involved:**
- `scripts/unpack_docx.py` → calls `ooxml_scripts/unpack.py`

**Output:**
```
unpacked/template/
├── _rels/
├── word/
│   ├── document.xml          ← Contains all text and placeholders
│   ├── styles.xml            ← Formatting definitions
│   ├── numbering.xml
│   ├── media/                ← Images, embedded files
│   └── ...
└── [Content_Types].xml
```

### Stage 2: LLM ANALYSIS - Understand & Map

**What happens:**
- Agent reads unpacked `word/document.xml`
- Agent extracts all `{{placeholder_name}}` patterns
- Agent analyzes template structure
- Agent creates `replacements.json` mapping placeholder→value

**Tools involved:**
- `tools.find_placeholders()` - Extracts all placeholder names
- `tools.create_replacement_mapping()` - Agent writes mapping

**Output:**
```json
{
  "employee_name": "John Smith",
  "employee_id": "EMP12345",
  "start_date": "2025-11-14",
  "salary": "$95,000"
}
```

### Stage 3: AUTOMATION - Fill & Pack

**What happens:**
- Script reads `replacements.json`
- Script finds all `{{placeholder}}` in XML text nodes
- Script replaces with actual values
- XML is validated and condensed
- Directory is repackaged into DOCX ZIP

**Scripts involved:**
- `scripts/fill_docx.py` - Performs replacements
- `scripts/pack_docx.py` → calls `ooxml_scripts/pack.py`

**Important**: Only text content is modified; all formatting, styles, and structure remain intact.

## Placeholder Syntax

In your DOCX template, use:

```
{{placeholder_name}}
```

**Rules:**
- Placeholder name is case-sensitive
- Only alphanumeric and underscore characters allowed: `[a-zA-Z0-9_]`
- Will be replaced with corresponding value from mapping
- Can appear anywhere in document text

**Examples:**
```
Hello {{first_name}} {{last_name}}!
Your ID: {{employee_id}}
Salary: {{salary}}
Date: {{hire_date}}
```

## Why Formatting is Preserved

The key to preserving formatting is the **unpack→modify→pack** workflow:

1. **Unpack** extracts entire DOCX with ALL XML files (styles, formatting, relationships)
2. **Modify** only changes text content in `word/document.xml` text nodes
3. **Pack** reconstructs DOCX with all original files intact

This means:
- ✅ Font sizes, colors, styles preserved
- ✅ Tables, images, headers/footers preserved
- ✅ Numbering, bullet points preserved
- ✅ Page layouts, margins preserved
- ✅ Only text values change

## Session Isolation

Each user session has completely isolated workspaces:

```
workspaces/
├── user_alice/
│   ├── input/          (Alice's uploaded files)
│   ├── unpacked/       (Alice's working directory)
│   ├── debug/          (Alice's analysis files)
│   └── output/         (Alice's filled documents)
│
└── user_bob/
    ├── input/          (Bob's uploaded files)
    ├── unpacked/       (Bob's working directory)
    ├── debug/          (Bob's analysis files)
    └── output/         (Bob's filled documents)
```

**Benefits:**
- Multiple users can work simultaneously
- No file interference between sessions
- Easy cleanup: `SessionTools.cleanup()`
- Perfect for shared/multi-user systems

## Integration with Agno Agent Framework

This codebase is designed to integrate with Agno AgentOS:

```python
from agno import Agent
from agents import SessionTools

# Agno agent calls these tools
def docx_agent():
    agent = Agent(
        name="DOCX Autofill",
        model="claude-sonnet-4.5",
        tools=[
            # Register session-aware tools
            unpack_template_tool,
            find_placeholders_tool,
            create_mapping_tool,
            fill_and_pack_tool,
            # ... etc
        ]
    )
```

## Example Workflow

### Simple Form Filling

```
User: "I have a template.docx with {{name}}, {{email}}, and {{date}}.
       Please fill it with: name='John Smith', email='john@example.com', date='2025-11-14'"

Agent:
1. Unpacks template.docx using scripts/unpack_docx.py
   → Creates unpacked/template/ directory with all XML

2. Finds placeholders using scripts/extract_placeholders.py
   → Returns: ["name", "email", "date"]

3. Creates replacement mapping
   → Saves to debug/replacements.json

4. Shows preview to user
   → Confirms: name → John Smith, email → john@example.com, date → 2025-11-14

5. Fills template using scripts/fill_docx.py
   → Replaces {{name}}, {{email}}, {{date}} in word/document.xml

6. Packs result using scripts/pack_docx.py
   → Creates output/template_filled.docx with all formatting intact
```

## Technical Implementation Details

### Unpack/Pack Pattern

The solution uses the **OOXML Unpack/Pack Pattern** from document-skills:

```
DOCX (ZIP Archive)
        ↓ unpack.py
     XML Files (pretty-printed)
        ↓ [modify text nodes]
     XML Files (modified)
        ↓ pack.py
   DOCX (ZIP Archive)
```

### XML Manipulation

The `fill_docx.py` script uses Python's `xml.etree.ElementTree` to:
- Parse `word/document.xml`
- Find all `<w:t>` (text) nodes
- Replace `{{placeholder}}` with actual values
- Preserve all XML structure and attributes

### Automation vs Analysis

| Stage | Type | Who/What | Output |
|-------|------|---------|--------|
| 1 | AUTOMATION | Script runs unpack.py | unpacked/template/ |
| 2 | LLM ANALYSIS | Agent reads XML, creates mapping | debug/replacements.json |
| 3 | AUTOMATION | Scripts run fill_docx.py + pack.py | output/filled.docx |

## File Structure Details

### ooxml_scripts/ (Copied from document-skills)

Contains the core unpack/pack utilities:
- `unpack.py` - Extracts DOCX ZIP and formats XML
- `pack.py` - Condenses XML and creates DOCX ZIP
- `validate.py` - Optional schema validation
- `validation/` - XSD schema validators

### lib/ (Copied from document-skills)

Contains the Document Library:
- `utilities.py` - XMLEditor class for DOM manipulation
- `document.py` - High-level Document API

### templates/ (Copied from document-skills)

Pre-made XML template files for:
- Comments and comment replies
- Extended/extensible comments
- People tracking for tracked changes

These are used by `document.py` for advanced features.

## Requirements

### Core Dependencies
- `defusedxml` - Secure XML parsing (minimal package)
- Python 3.7+

### Optional
- `lxml` - For schema validation (included in validation modules if used)
- `soffice` (LibreOffice) - For document validation (optional, pack.py works without it with `--force`)

## Standalone Nature

This codebase is **completely standalone**:
- ✅ All necessary scripts are included locally
- ✅ No external calls to document-skills directory
- ✅ No complex dependencies
- ✅ Can be deployed as a complete unit
- ✅ Works offline (no network calls needed)

## Security Considerations

- **XML Safety**: Uses `defusedxml` for secure XML parsing (prevents billion laughs attacks)
- **Path Validation**: Filenames sanitized, no directory traversal
- **Session Isolation**: Cannot access other sessions' files
- **No Code Execution**: Only XML manipulation, no eval/exec

## Troubleshooting

### "document.xml not found"
- Ensure DOCX file is valid (open in Word/LibreOffice)
- Check file hasn't been corrupted

### "Placeholders not found"
- Verify template uses `{{placeholder}}` syntax
- Check placeholder names are spelled correctly
- Ensure XML wasn't corrupted during unpacking

### "File not found in session"
- Use `SessionTools.list_input_files()` to see available files
- Check filename spelling
- Verify file was uploaded to same session

## Next Steps for Enhancement

The current implementation handles basic autofill. Potential enhancements:

- Tracked changes (mark filled cells with insertions)
- Comments (add comments to filled fields)
- Batch filling (fill multiple templates with CSV data)
- Template validation (check template is valid before filling)
- Data extraction (auto-extract data from source DOCX)

## License

Based on document-skills framework. Use according to document-skills LICENSE.txt terms.

---

**Built for intelligent document automation with Claude** 📄✨
