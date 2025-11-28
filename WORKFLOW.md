# DOCX Autofill Agent - Workflow Documentation

## Overview

The DOCX Autofill Agent intelligently fills DOCX templates with data extracted from source documents. It assesses template complexity and automatically selects the best filling approach:
- **LOW Complexity** (simple templates) → Python fill_fields()
- **HIGH Complexity** (professional templates) → Node.js docx-js library

---

## 6-Phase Workflow

```
Phase 1: PREPARATION & SETUP
  ├─ Unpack template to XML
  ├─ Convert documents to markdown
  ├─ Assess complexity (LOW vs HIGH) ← KEY DECISION POINT
  └─ Insert placeholders if needed

Phase 2: DATA EXTRACTION
  ├─ Extract text content
  ├─ Extract tables
  └─ Extract SDT form fields

Phase 3 & 4: COMPLEXITY-BASED FILLING
  ├─ LOW: fill_fields() - Multi-strategy placeholder replacement
  └─ HIGH: docx-js - Fresh document generation + headers/footers/branding preservation

Phase 5: VALIDATION (AUTOMATIC)
  ├─ Check all placeholders filled
  ├─ Verify document structure
  └─ Validate XML well-formedness

Phase 6: OUTPUT & PACKAGING
  └─ Pack filled XML back to DOCX format
```

---

## Phase Details

### Phase 1: Preparation & Setup

**What happens:**
- Template DOCX unpacked to XML structure
- Documents converted to markdown for analysis
- **Complexity Assessment** (Phase 1.3):
  - **LOW**: Simple forms, letters, basic invoices → Use fill_fields()
  - **HIGH**: Multi-section reports, charters, branded documents → Use docx-js

**Tools:** `unpack_template()`, `convert_docx_to_markdown()`, `insert_placeholders()`

---

### Phase 2: Data Extraction

**What happens:**
- Data extracted from source using 3 methods:
  1. Text extraction (paragraphs, runs)
  2. Table extraction (structured data)
  3. SDT form field extraction
- Results merged into `extracted_values` dictionary

**Tools:** `extract_all_data()`

**Output:** `debug/extraction_results.json` with all extracted data

---

### Phase 3 & 4: Template Filling (Complexity-Based)

#### LOW Complexity Approach

**Process:**
1. Create field mapping from extracted data
2. Call `fill_fields(field_mapping)`
3. 5 filling strategies:
   - Text placeholder replacement (`{{FIELD}}`)
   - SDT form field matching
   - Element ID-based targeting
   - Multi-run placeholder handling
   - Table row filling

**Professional Filling Guidelines:**
- ✅ Fill all sections that need data
- ✅ Keep original formatting intact
- ✅ Maintain professional appearance
- ❌ DO NOT add/remove/modify sections (anti-hallucination)

**Tools:** `fill_fields()`

---

#### HIGH Complexity Approach

**Process:**
1. Read `lib/docx-js.md` for docx library reference
2. Create JavaScript code that builds fresh document using docx library
3. Populate with extracted data in proper structure
4. Save JS file using `save_debug_file()`
5. Execute via `run_node_script()` (Node.js subprocess)
6. Preserve original headers, footers, branding:
   - Unpack both original and generated documents
   - Copy header*.xml, footer*.xml files
   - Copy word/media/ folder (logos, images)
   - Update relationships and content types
   - Ensure image references (rIds) are correct
   - Repack final DOCX

**Tools:** `read_lib_file()`, `save_debug_file()`, `run_node_script()`

**Result:** Professional document with fresh content + original branding

---

### Phase 5: Validation (Automatic)

**3-Tier Validation:**
1. **Placeholders**: Verify all `{{FIELD}}` patterns removed
2. **Structure**: Verify document XML structure intact
3. **Well-formedness**: Verify XML is valid

Automatic during Phase 3. Ensures safe DOCX generation.

---

### Phase 6: Output & Packaging

**Process:**
- Pack filled XML back to DOCX format
- Save to `output/` directory
- Ready for download

**Tools:** `pack_template()`

---

## Architecture

### 3-Layer Design

```
Layer 1: Core Logic (docx_tools.py)
  ├─ extract_all_data()
  ├─ fill_fields()
  ├─ insert_placeholders()
  ├─ pack_template()
  └─ All XML manipulation and filling logic

Layer 2: Tool Functions (docx_session_aware_tool_functions.py)
  └─ Wrapper functions for agent tool registration
  └─ Converts Layer 1 results to JSON for agent

Layer 3: Session Manager (docx_session_tools.py)
  └─ Workspace isolation per user/session
  └─ Maps relative paths to session directories
  └─ Manages session-specific file operations

Agent (docx_agent_with_sessions.py)
  ├─ Registers 18 tools
  ├─ Implements workflow instructions
  ├─ Makes intelligent decisions
  ├─ Custom DocxJsTools for Node.js execution
  └─ Complexity-based decision logic
```

---

## 18 Registered Tools

**File Discovery (3):**
- `list_input_files()` - See uploaded files
- `list_output_files()` - See generated files
- `get_session_info()` - Session workspace info

**Phase 1 (3):**
- `unpack_template(filename)` - Extract DOCX to XML
- `convert_docx_to_markdown(filename)` - Convert to markdown
- `insert_placeholders(field_analysis)` - Auto-insert {{FIELD}} placeholders

**Phase 2 (3):**
- `extract_all_data(filename)` - Extract using 3 methods
- `read_text_file(filename)` - Read markdown/text
- `read_json_file(filename)` - Read JSON debug files

**Phase 3/4 - Filling (2):**
- `fill_fields(field_mapping)` - LOW complexity filling
- `pack_template(filename)` - Pack to DOCX

**Phase 4 - docx-js HIGH Complexity (3):**
- `read_lib_file(filename)` - Read lib/docx-js.md
- `save_debug_file(filename, content)` - Save JS to disk
- `run_node_script(script, output_path)` - Execute Node.js

**Utilities (3):**
- `cleanup()` - Delete session workspace
- `PythonTools()` - Complex Python analysis
- `convert_docx_to_markdown()` - Markdown preview

---

## Session Management

Each user/session has isolated workspace:

```
workspaces/{session_id}/
├── input/           (uploaded files)
├── unpacked/        (extracted XML)
│   └── template/    (Phase 1: unpacked structure)
├── debug/           (intermediate files)
│   ├── template.md
│   ├── extraction_results.json
│   ├── field_mapping_applied.json
│   └── filling_results.json
└── output/          (Phase 6: final DOCX)
    └── filled.docx
```

Benefits:
- Multiple concurrent users
- No file interference
- Complete isolation
- Auto-cleanup available

---

## Setup & Prerequisites

### Required

1. **Python 3.8+** with packages:
   - python-docx
   - lxml
   - pandoc (system)

2. **Node.js 14+** with npm packages:
   ```bash
   npm install docx
   ```

3. **Pandoc** (for markdown conversion):
   ```bash
   # Linux
   sudo apt-get install pandoc

   # macOS
   brew install pandoc

   # Windows
   choco install pandoc
   ```

### Key Components

- **Subprocess Execution**: Python subprocess.run() calls Node.js for HIGH complexity documents
- **docx-js Library**: Node.js library for professional DOCX generation
- **Custom Toolkit**: DocxJsTools for reliable Node.js execution

---

## Quick Start for Developers

### Understanding the Decision Flow

1. **Read Phase 1.3 Assessment** in agent instructions (docx_agent_with_sessions.py lines 115-153)
   - How complexity is determined
   - Decision criteria for LOW vs HIGH

2. **Understand Two Approaches**
   - LOW: Simple XML manipulation via fill_fields()
   - HIGH: Fresh document generation via docx-js + header/footer preservation

3. **Key Files to Review**
   - `agents/docx_tools.py` - Core business logic
   - `agents/docx_session_tools.py` - Session/workspace management
   - `lib/docx-js.md` - docx library reference (read by agent for HIGH complexity)
   - `agents/docx_agent_with_sessions.py` - Agent instructions and tool registration

### Common Development Tasks

**To add a new extraction method:**
1. Modify `extract_all_data()` in docx_tools.py
2. Add new extraction method (similar to text/table/SDT)
3. Merge results into extracted_values

**To add a new filling strategy:**
1. Modify `fill_fields()` in docx_tools.py
2. Add new strategy (similar to A-E strategies)
3. Update results with strategy_used tracking

**To modify complexity assessment:**
1. Edit Phase 1.3 complexity assessment logic in agent instructions
2. Change decision criteria (LOW vs HIGH)
3. Test with different template types

---

## Important Notes

### Anti-Hallucination Safeguards

Agent instructions explicitly prevent:
- ❌ Adding new sections not in original template
- ❌ Removing sections from template
- ❌ Modifying section structure
- ✅ Filling existing sections with data
- ✅ Leaving empty sections blank if no data

This ensures filled documents match original structure exactly.

### Headers/Footers Preservation (HIGH Complexity)

When using docx-js approach:
1. Original template headers/footers preserved separately
2. After fresh document generation, headers/footers copied back
3. All relationships updated to point to correct media files
4. Company logos and branding maintained

This is critical for professional documents.

### Subprocess & Node.js Integration

For HIGH complexity documents:
- Python subprocess.run() executes Node.js
- JavaScript file created with save_debug_file()
- Node.js generates DOCX using docx library
- Packer.toBuffer() + fs.writeFileSync() saves file
- Return path used for header/footer preservation

---

## Troubleshooting

**"Error reading file: No such file or directory: 'D:\lib\docx-js.md'"**
- Agent must use `read_lib_file("docx-js.md")` NOT PythonTools
- This tool resolves correct file path dynamically

**"File not found when running node script"**
- JavaScript must be saved to disk BEFORE execution
- Use `save_debug_file()` which returns full path
- Pass returned path to `run_node_script()`

**Headers/footers not showing in output**
- Check relationship IDs are updated correctly
- Verify media/ folder copied with correct rId mappings
- Ensure [Content_Types].xml has header/footer entries

**Document filled but looks awkward**
- Check complexity was assessed correctly
- LOW complexity should look fine
- HIGH complexity should look professional with docx-js
- Review Professional Filling Guidelines in agent instructions

---

## Summary

The DOCX Autofill Agent is a **6-phase intelligent system** that:

1. ✅ Assesses template complexity automatically
2. ✅ Selects best filling approach (LOW or HIGH)
3. ✅ Extracts data using 3 methods
4. ✅ Fills templates with 5+ strategies or fresh document generation
5. ✅ Validates at 3 tiers
6. ✅ Produces professional output with formatting preserved

**Key Innovation**: Complexity-based decision logic provides both speed (LOW) and quality (HIGH).

**Key Architecture**: 3-layer design with session isolation enables multi-user concurrent processing.

**Key Tools**: 18 registered tools handle all workflow phases with workspace awareness.

For questions about agent behavior, see `agents/docx_agent_with_sessions.py` lines 100-420 for complete workflow instructions.
