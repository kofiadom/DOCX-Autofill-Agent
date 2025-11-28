# DOCX Autofill Agent

An intelligent Agno agent that fills DOCX templates with extracted data. Uses Claude AI to assess template complexity and automatically selects the best filling approach.

**Status:** ✅ Production Ready
**Version:** 1.0
**Python:** 3.8+ | **Node.js:** 14+

---

## Quick Overview

```
Input: template.docx + source.docx (or manual mapping)
         ↓
    Agent reads template & source
         ↓
    Assess complexity: Simple or Professional?
         ↓
    LOW Complexity  →  Python fill_fields()     → Fast, XML-based
    HIGH Complexity →  Node.js docx-js library  → Professional quality
         ↓
    Output: filled.docx (with all formatting preserved)
```

---

## 6-Phase Workflow

The agent follows a structured 6-phase process:

| Phase | What | How |
|-------|------|-----|
| **1** | Preparation & Setup | Unpack template, assess complexity (LOW vs HIGH) |
| **2** | Data Extraction | Extract text, tables, form fields from source |
| **3&4** | Complexity-Based Filling | LOW: fill_fields() / HIGH: docx-js |
| **5** | Validation | 3-tier validation (placeholders, structure, XML) |
| **6** | Output & Packaging | Pack to DOCX format |

See **WORKFLOW.md** for detailed phase descriptions.

---

## Two Filling Approaches

### LOW Complexity Templates
**Use case:** Forms, simple letters, basic invoices
- **Method:** Python `fill_fields()` with multi-strategy approach
- **Speed:** Fast (1-3 seconds)
- **Formatting:** All preserved
- **Best for:** Simple templates with straightforward placeholders

### HIGH Complexity Templates
**Use case:** Professional reports, branded documents, multi-section charters
- **Method:** Node.js `docx-js` library for fresh document generation
- **Speed:** Slightly slower (3-5 seconds) but produces professional quality
- **Formatting:** Original headers/footers/branding preserved
- **Best for:** Complex layouts with multiple sections, headers/footers, company branding

---

## Installation

### Prerequisites

```bash
# Python 3.8+
python --version

# Node.js 14+
node --version
npm --version
```

### Setup

```bash
# 1. Install Python dependencies
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# 2. Install Node.js docx library
npm install docx

# 3. Install Pandoc (for markdown conversion)
# Linux
sudo apt-get install pandoc
# macOS
brew install pandoc
# Windows
choco install pandoc
```

---

## Architecture

### 3-Layer Design

```
┌─────────────────────────────────────┐
│  Agent (docx_agent_with_sessions.py) │
│  - Workflow instructions             │
│  - Complexity-based decisions        │
│  - 18 registered tools               │
└──────────────────┬──────────────────┘
                   │
┌──────────────────┴──────────────────┐
│ Layer 2: Tool Functions             │
│ (docx_session_aware_tool_functions) │
│ - Wrapper functions for agent       │
│ - Converts results to JSON          │
└──────────────────┬──────────────────┘
                   │
┌──────────────────┴──────────────────┐
│ Layer 3: Session Manager            │
│ (docx_session_tools)                │
│ - Workspace isolation               │
│ - Path resolution                   │
│ - File operations                   │
└──────────────────┬──────────────────┘
                   │
┌──────────────────┴──────────────────┐
│ Layer 1: Core Logic                 │
│ (docx_tools)                        │
│ - XML manipulation                  │
│ - Data extraction                   │
│ - Field filling                     │
│ - Document packing                  │
└─────────────────────────────────────┘
```

### Session Isolation

Each user has isolated workspace:

```
workspaces/{session_id}/
├── input/       (uploaded files)
├── unpacked/    (extracted XML)
├── debug/       (intermediate files: extraction, field mapping, results)
└── output/      (filled.docx)
```

---

## 18 Registered Tools

### File Discovery (3)
- `list_input_files()` - See uploaded files
- `list_output_files()` - See generated files
- `get_session_info()` - Session workspace info

### Phase 1: Preparation (3)
- `unpack_template(filename)` - Extract DOCX to XML
- `convert_docx_to_markdown(filename)` - Convert to markdown
- `insert_placeholders(field_analysis)` - Auto-insert {{FIELD}} placeholders

### Phase 2: Data Extraction (3)
- `extract_all_data(filename)` - Extract text, tables, SDT fields
- `read_text_file(filename)` - Read markdown/text files
- `read_json_file(filename)` - Read debug JSON files

### Phase 3/4: Filling (2)
- `fill_fields(field_mapping)` - LOW complexity filling
- `pack_template(filename)` - Pack to DOCX

### Phase 4: High Complexity docx-js (3)
- `read_lib_file(filename)` - Read lib/docx-js.md reference
- `save_debug_file(filename, content)` - Save JS to disk
- `run_node_script(script, output_path)` - Execute Node.js subprocess

### Utilities (3)
- `cleanup()` - Delete session workspace
- `PythonTools()` - Complex Python analysis
- `convert_docx_to_markdown()` - Markdown preview

---

## Usage

### For Users

1. **Upload template.docx** and optionally **source.docx**
2. **Agent analyzes** and assesses complexity
3. **Agent extracts** data from source (if provided)
4. **Agent fills** template using appropriate approach
5. **Download filled.docx**

### For Developers

See **WORKFLOW.md** for:
- Detailed phase descriptions
- Architecture overview
- Quick start guide
- Development tasks

Key files to review:
- `agents/docx_tools.py` - Core business logic
- `agents/docx_session_tools.py` - Session management
- `agents/docx_agent_with_sessions.py` - Agent instructions & tools
- `lib/docx-js.md` - docx library reference (used by agent)

---

## Key Features

✅ **Intelligent Complexity Assessment** - Automatically selects approach
✅ **Two Filling Methods** - Speed (LOW) or Quality (HIGH)
✅ **Complete Session Isolation** - Multi-user safe
✅ **Comprehensive Data Extraction** - Text, tables, form fields
✅ **Professional Output** - Headers/footers/branding preserved
✅ **3-Tier Validation** - Ensures document integrity
✅ **Anti-Hallucination** - Strict section structure adherence
✅ **Debug Files** - Full transparency at each phase

---

## How It Works

### Complexity Assessment (Phase 1)

Agent reads template markdown and assesses:

- **LOW:** Single section, basic placeholders, minimal formatting
  - ✅ Use Python fill_fields()

- **HIGH:** Multiple sections, tables, headers/footers, complex formatting
  - ✅ Use docx-js for fresh professional document

### Data Extraction (Phase 2)

Extracts from source using 3 methods:
1. **Text extraction** - Paragraphs and runs
2. **Table extraction** - Structured data
3. **SDT extraction** - Form fields

Results merged into `extracted_values` for mapping.

### Template Filling (Phase 3 & 4)

**LOW Complexity:**
- Multi-strategy placeholder replacement
- 5 different filling strategies for maximum coverage
- Preserves all original formatting

**HIGH Complexity:**
- Generates fresh document using docx-js library
- Copies headers, footers, media from original
- Updates all relationships for proper display
- Produces professional-quality output

### Validation (Phase 5)

Automatic 3-tier validation:
1. **Placeholder Check** - All {{FIELD}} patterns removed
2. **Structure Check** - Document XML structure intact
3. **Well-Formedness** - XML is valid

### Output (Phase 6)

Packs filled XML back to DOCX format, ready for download.

---

## Important Implementation Details

### Subprocess & Node.js Integration

For HIGH complexity templates:
```
Agent creates JavaScript code
    ↓
save_debug_file() → saves to disk
    ↓
run_node_script() → Python subprocess.run(["node", script, output_path])
    ↓
Node.js executes, generates DOCX
    ↓
JavaScript saves via fs.writeFileSync()
    ↓
Headers/footers/branding copied from original
    ↓
Final DOCX ready
```

### Anti-Hallucination Safeguards

Agent instructions explicitly prevent:
- ❌ Adding sections not in original template
- ❌ Removing sections
- ❌ Modifying section structure
- ✅ Filling existing sections with data
- ✅ Leaving empty sections blank

### Professional Filling Guidelines

When filling templates:
- ✅ Fill all sections that need data
- ✅ Keep original formatting intact
- ✅ Maintain professional appearance
- ✅ Preserve authenticity (if quality not compromised)
- ❌ DO NOT hallucinate content

---

## Troubleshooting

**"Error reading file: docx-js.md"**
- Agent must use `read_lib_file()` NOT PythonTools
- This tool resolves correct file path dynamically

**"File not found when running node script"**
- JavaScript must be saved to disk BEFORE execution
- Use `save_debug_file()` which returns full path
- Pass returned path to `run_node_script()`

**"Headers/footers not showing"**
- Check relationship IDs are updated correctly
- Verify media/ folder copied with correct rId mappings
- Ensure [Content_Types].xml has header/footer entries

**"Template filled but looks awkward"**
- Check complexity was assessed correctly
- LOW complexity should look fine
- HIGH complexity should be professional with docx-js
- Review Professional Filling Guidelines in agent instructions

---

## Project Structure

```
Agno-DOCX-Autofill-Agent/
├── agents/
│   ├── docx_agent_with_sessions.py     (Agent + tool registration)
│   ├── docx_session_aware_tool_functions.py  (Layer 2: Tool functions)
│   ├── docx_session_tools.py           (Layer 3: Session manager)
│   ├── docx_tools.py                   (Layer 1: Core logic)
│   ├── extraction_module.py
│   ├── filling_strategies.py
│   └── validation_module.py
├── lib/
│   ├── docx-js.md                      (docx library reference for agent)
│   └── __init__.py
├── scripts/
│   └── pack_docx.py                    (Pack utility)
├── ooxml_scripts/
│   └── validation/                     (XML validation)
├── workspaces/                         (Session directories - auto-created)
├── package.json                        (npm: docx library)
├── package-lock.json
├── WORKFLOW.md                         (Complete workflow documentation)
├── requirements.txt
└── README.md
```

---

## Requirements

### Python Packages
- python-docx
- lxml
- defusedxml

### System Requirements
- Pandoc (for markdown conversion)
- Python 3.8+
- Node.js 14+

### npm Packages
- docx

---

## Key Differences from Older Approaches

| Aspect | Old | New |
|--------|-----|-----|
| **Phases** | 3 stages | 6 phases |
| **Complexity** | Single approach | LOW vs HIGH |
| **Quality** | Basic XML filling | Professional docx-js |
| **Headers/Footers** | Not preserved | Preserved (HIGH) |
| **Speed** | Fast but limited | Optimized per complexity |
| **Output Quality** | Consistent | Professional (HIGH) |

---

## For New Developers

1. **Start with:** WORKFLOW.md (375 lines, quick read)
2. **Understand:** 6-phase workflow and complexity decision
3. **Review:** Agent instructions in docx_agent_with_sessions.py (lines 100-420)
4. **Explore:** docx_tools.py for core logic
5. **Add features:** Follow pattern in Common Development Tasks in WORKFLOW.md

---

## Summary

**DOCX Autofill Agent** is an intelligent system that:

1. ✅ Automatically assesses template complexity
2. ✅ Selects optimal filling approach
3. ✅ Extracts data using 3 methods
4. ✅ Fills templates with professional quality
5. ✅ Validates output integrity
6. ✅ Produces perfect DOCX files

**Innovation:** Complexity-based decision logic provides both speed and quality.

**Architecture:** 3-layer design enables scalable multi-user processing.

**Key Tech:** Python + Node.js + Agno Agent Framework.

---

**Built for intelligent document automation** 📄✨
