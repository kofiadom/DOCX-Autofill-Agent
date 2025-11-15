# DOCX Autofill Agent - Complete Workflow Documentation

## Table of Contents

1. [Quick Overview](#quick-overview)
2. [Three-Stage Architecture](#three-stage-architecture)
3. [Detailed Workflow Breakdown](#detailed-workflow-breakdown)
4. [File Flow Diagram](#file-flow-diagram)
5. [Usage Examples](#usage-examples)
6. [Session Management](#session-management)
7. [Tool Reference](#tool-reference)
8. [Data Structures](#data-structures)
9. [Troubleshooting](#troubleshooting)

---

## Quick Overview

The DOCX Autofill Agent intelligently fills DOCX templates using data from source documents. It follows a **three-stage workflow**:

```
┌─────────────────────────────────────────────────────────────────┐
│                     DOCX Autofill Workflow                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  STAGE 1: AUTOMATION                STAGE 2: LLM ANALYSIS        │
│  (Scripts & Automation)             (Agent Intelligence)         │
│                                                                 │
│  • Unpack DOCX → XML                • Read source markdown       │
│  • Extract {{placeholders}}         • Analyze requirements       │
│  • Convert source to markdown       • Create intelligent mapping │
│                                                                 │
│                    STAGE 3: AUTOMATION                          │
│                    (Scripts & Automation)                       │
│                                                                 │
│                  • Fill placeholders with values                │
│                  • Pack XML → DOCX                              │
│                  • Return filled document                       │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

**Key Principle**: Agent handles intelligent analysis; automation handles file operations.

---

## Three-Stage Architecture

### Stage 1: AUTOMATION - Unpack & Prepare

**What happens**: Scripts extract template structure and convert documents to readable format.

**Steps**:
1. User uploads `template.docx` and optionally `source.docx` to session input directory
2. Agent calls `unpack_template("template.docx")`
3. SessionAwareDocxTools maps to: `workspaces/{session_id}/input/template.docx`
4. Script calls `ooxml_scripts/unpack.py` which:
   - Extracts DOCX (ZIP archive) to XML directory
   - Preserves all formatting, styles, relationships
   - Creates structure: `word/document.xml`, `word/styles.xml`, etc.
5. If source document provided:
   - Agent calls `convert_docx_to_markdown("source.docx")`
   - Calls pandoc to convert DOCX to markdown
   - Preserves document structure and hierarchy
   - Saves to: `workspaces/{session_id}/debug/source.md`
6. Agent also converts template:
   - Agent calls `convert_docx_to_markdown("template.docx")`
   - Saves to: `workspaces/{session_id}/debug/template.md`

**Output of Stage 1**:
- Unpacked template XML: `workspaces/{session_id}/unpacked/template/`
- Template markdown: `workspaces/{session_id}/debug/template.md`
- Source markdown: `workspaces/{session_id}/debug/source.md` (if source provided)

**Files Used**:
- `ooxml_scripts/unpack.py` (extracts DOCX to XML)
- `scripts/docx_to_markdown.py` (calls pandoc for conversion)

---

### Stage 2: LLM ANALYSIS - Understand & Map

**What happens**: Agent applies semantic analysis to identify placeholders and create field mappings.

**Steps**:
1. Agent reads template markdown for semantic analysis:
   ```python
   template_content = read_text_file("template.md")
   ```
2. Agent analyzes template to identify ALL placeholder types:
   - **Explicit markers**: `{{name}}`, `[Field]`, `__field__`, `<<text>>`
   - **Implicit markers**: Lines of underscores, dots, spaces (visual empty fields)
   - **Placeholder hints**: Text with "fill in", "enter", "provide", "specify"
   - **Semantic gaps**: Labels without values ("Date:", "Signature:", etc.)
3. If source document provided, Agent reads source markdown:
   ```python
   source_content = read_text_file("source.md")
   ```
4. Agent extracts values from source using 4-level matching:
   - **Level 1**: Exact name match (case-insensitive)
   - **Level 2**: Fuzzy name match (variations like "date" → "Start Date")
   - **Level 3**: Context matching (financial values, date patterns, etc.)
   - **Level 4**: Ask user (if no clear match found)
5. Agent handles format conversions:
   - Dates: "November 14, 2025" → "11/14/2025"
   - Currency: "50000" → "$50,000"
   - Names: "JOHN SMITH" → "John Smith"
   - Numbers: "50000.00" → "50,000"
6. Agent creates field mapping dictionary:
   ```python
   {
     "Project Manager": "John Smith",
     "Date": "11/14/2025",
     "Budget": "$50,000"
   }
   ```
7. Agent shows analysis to user for confirmation:
   - Lists identified fields and extracted values
   - Asks: "Is this correct? Any changes needed?"
   - User can approve, adjust, or ask questions

**Agent Responsibilities**:
- Semantic analysis of template structure (not pattern matching)
- Intelligent field identification with confidence levels
- Semantic analysis of source document
- Smart data extraction and matching using 4-level strategy
- Format conversion and normalization
- Transparent analysis presentation to user
- Handling ambiguous cases with user clarification

**Output of Stage 2**:
- Field mapping: `field_mapping = {"label": "value", ...}`
- Preview shown to user for confirmation
- Ready for Stage 3 after user approval

**Tools Used**:
- `read_text_file()` - Read markdown files for analysis

---

### Stage 3: AUTOMATION - Fill & Pack

**What happens**: fill_fields tool applies field mapping and pack_template creates final DOCX.

**Three-Layer Tool Architecture**:

```
Layer 2: Agent Tool Function
  agents/docx_session_aware_tool_functions.py
  fill_fields(run_context, field_mapping)
         ↓
Layer 3: Session Wrapper
  agents/docx_session_tools.py
  SessionAwareDocxTools.fill_fields(field_mapping)
         ↓
Layer 1: Core Business Logic
  agents/docx_tools.py
  fill_fields(unpacked_dir, field_mapping)
         ↓
Document Library
  lib/document.py
  Document class (uses ooxml_scripts)
```

**Steps**:
1. After user approves analysis, agent calls `fill_fields(field_mapping)`:
   - Passes mapping dict like: `{"Project Manager": "John Smith", "Date": "11/14/2025"}`
   - Tool locates unpacked template directory
   - Imports Document Library from lib/document.py
   - For each label→value in field_mapping:
     - Finds paragraph containing label text
     - Searches forward for first empty `<w:t>` (text) element
     - Fills empty element with value
     - Document Library preserves all formatting automatically
   - Saves modified XML back to unpacked directory
   - Returns status: "✅ Filled 3 field(s): Project Manager, Date, Budget"

2. Agent calls `pack_template("filled.docx")`:
   - Calls ooxml_scripts/pack.py
   - Re-packages XML directory back to DOCX
   - Creates valid OOXML structure with proper relationships
   - Saves to: `workspaces/{session_id}/output/filled.docx`

3. User downloads filled document

**Why Formatting Is Always Preserved**:
- Document Library handles all XML infrastructure
- Only text content (`<w:t>` elements) changes
- XML structure (`<w:r>`, `<w:p>`, styles) completely preserved
- All formatting attributes (bold, italic, colors, sizes) unchanged
- Namespaces, RSID values, relationships all maintained
- No manual XML manipulation required

**Output of Stage 3**:
- Filled DOCX: `workspaces/{session_id}/output/filled.docx`

**Tools Used**:
- `fill_fields(field_mapping)` - Fills fields using Document Library (safe, automatic)
- `pack_template(filename)` - Packs XML back to DOCX

---

## Detailed Workflow Breakdown

### Complete Workflow: Fill Template with Source Data

```
1. USER INTERACTION
   └─ Upload files to session:
      - template.docx
      - source.docx
      └─ REST API POST /api/upload/{session_id}

2. AGENT STARTS PROCESSING
   └─ Agent receives user message with session context

3. STAGE 1: AUTOMATION - Unpack & Extract
   ├─ Agent calls: unpack_template("template.docx")
   │  ├─ SessionAwareDocxTools.unpack_template()
   │  │  ├─ Maps relative filename to: workspaces/{session_id}/input/template.docx
   │  │  └─ Calls docx_tools.unpack_docx()
   │  │     ├─ Subprocess runs: scripts/unpack_docx.py
   │  │     └─ Which calls: ooxml_scripts/unpack.py
   │  │        ├─ Extracts DOCX (ZIP)
   │  │        └─ Creates: workspaces/{session_id}/unpacked/template/
   │  │           ├─ word/document.xml
   │  │           ├─ word/styles.xml
   │  │           ├─ word/relationships/document.xml.rels
   │  │           ├─ _rels/.rels
   │  │           └─ [Content_Types].xml
   │  └─ Returns: "✅ Template unpacked to working directory"
   │
   ├─ Agent calls: find_placeholders()
   │  ├─ SessionAwareDocxTools.find_placeholders()
   │  │  ├─ Finds unpacked directory
   │  │  └─ Calls docx_tools.extract_placeholders()
   │  │     └─ Subprocess runs: scripts/extract_placeholders.py
   │  │        ├─ Imports: XMLEditor from lib/utilities.py
   │  │        ├─ Parses: word/document.xml
   │  │        ├─ Uses: dom.getElementsByTagNameNS() for namespace-aware access
   │  │        └─ Extracts: {{placeholder}} patterns via regex
   │  └─ Returns: JSON {"placeholders": ["name", "company", "amount"], "count": 3}
   │
   └─ Agent calls: convert_docx_to_markdown("source.docx")
      ├─ SessionAwareDocxTools.convert_docx_to_markdown()
      │  ├─ Maps: source.docx → workspaces/{session_id}/input/source.docx
      │  └─ Calls docx_tools.convert_docx_to_markdown()
      │     └─ Subprocess runs: scripts/docx_to_markdown.py
      │        ├─ Calls: pandoc source.docx -o source.md
      │        └─ Saves to: workspaces/{session_id}/debug/source.md
      └─ Returns: "✅ Converted source.docx to markdown"

4. STAGE 2: LLM ANALYSIS - Understand & Map
   ├─ Agent calls: read_text_file("source.md")
   │  └─ Returns: Markdown content of source document
   │
   ├─ Agent ANALYZES:
   │  ├─ Reads placeholder list: ["name", "company", "amount", "date"]
   │  ├─ Reads source content
   │  ├─ Matches placeholders to source data:
   │  │  ├─ "name" → "John Smith" (from document)
   │  │  ├─ "company" → "Acme Corporation" (extracted from heading)
   │  │  ├─ "amount" → "$50,000" (from financial section)
   │  │  └─ "date" → "2025-11-14" (from header)
   │  └─ Creates mapping dictionary
   │
   ├─ Agent calls: create_replacement_mapping({...})
   │  ├─ SessionAwareDocxTools.create_replacement_mapping()
   │  │  ├─ Saves JSON to: workspaces/{session_id}/debug/replacements.json
   │  │  └─ Validates coverage
   │  └─ Returns: "✅ Created replacement mapping"
   │
   └─ Agent calls: generate_preview()
      ├─ SessionAwareDocxTools.generate_preview()
      │  ├─ Reads: workspaces/{session_id}/debug/replacements.json
      │  └─ Formats for display
      └─ Returns:
         ```
         {{name}} → John Smith
         {{company}} → Acme Corporation
         {{amount}} → $50,000
         {{date}} → 2025-11-14
         ```

5. AGENT ASKS USER FOR APPROVAL
   └─ Shows preview and asks: "Does this look correct?"

6. USER APPROVES
   └─ User responds: "Yes, looks good"

7. STAGE 3: AUTOMATION - Fill & Pack
   ├─ Agent calls: fill_template()
   │  ├─ SessionAwareDocxTools.fill_template()
   │  │  ├─ Reads: workspaces/{session_id}/debug/replacements.json
   │  │  └─ Calls docx_tools.fill_docx()
   │  │     └─ Subprocess runs: scripts/fill_docx.py
   │  │        ├─ Imports: Document from lib/document.py
   │  │        ├─ Opens: workspaces/{session_id}/unpacked/template/
   │  │        ├─ For each replacement:
   │  │        │  ├─ Finds: get_node(tag="w:r", contains="{{name}}")
   │  │        │  ├─ Creates: <w:r><w:t>John Smith</w:t></w:r>
   │  │        │  └─ Replaces: replace_node(old_node, new_xml)
   │  │        └─ Saves: Modified unpacked directory
   │  └─ Returns: "✅ Template filled with replacements"
   │
   └─ Agent calls: pack_template("filled.docx")
      ├─ SessionAwareDocxTools.pack_template()
      │  ├─ Calls docx_tools.pack_docx()
      │  │  └─ Subprocess runs: scripts/pack_docx.py
      │  │     ├─ Calls: ooxml_scripts/pack.py
      │  │     ├─ Re-packages: workspaces/{session_id}/unpacked/template/ → DOCX
      │  │     └─ Saves to: workspaces/{session_id}/output/filled.docx
      │  └─ Returns: "✅ Packed document"
      └─ Returns: File path info

8. DELIVERY
   └─ User downloads filled.docx from output directory
```

---

## File Flow Diagram

```
STAGE 1: UNPACK & EXTRACT
═════════════════════════════════════════════════════════════════

Input:
  workspaces/{session_id}/input/
  ├── template.docx
  └── source.docx
         ↓
    [unpack_docx.py]
    [extract_placeholders.py]
    [docx_to_markdown.py]
         ↓
Output:
  workspaces/{session_id}/
  ├── unpacked/template/
  │   ├── word/document.xml          ← Modified here in Stage 3
  │   ├── word/styles.xml
  │   ├── word/relationships/
  │   └── _rels/
  │
  └── debug/
      ├── source.md                  ← Read by agent in Stage 2
      ├── placeholders.json           ← List of found placeholders
      └── replacements.json           ← Created by agent in Stage 2


STAGE 2: ANALYZE & MAP
═════════════════════════════════════════════════════════════════

Read Files:
  ├── debug/source.md
  │   └─ Used to extract values
  │
  └── debug/placeholders.json
      └─ Determines what needs filling

Agent Processing:
  ├─ Parse source.md
  ├─ Match placeholders to source data
  ├─ Create intelligent mappings
  └─ Show preview to user

Write Files:
  └── debug/replacements.json
      ├─ Created by create_replacement_mapping()
      └─ Contains: {
           "name": "value",
           "company": "value",
           ...
         }


STAGE 3: FILL & PACK
═════════════════════════════════════════════════════════════════

Read Files:
  ├── unpacked/template/word/document.xml
  │   └─ Contains {{placeholders}}
  │
  └── debug/replacements.json
      └─ Contains values to fill

Processing:
  [fill_docx.py]
  ├─ Load Document from unpacked/template/
  ├─ For each {{placeholder}}:
  │  ├─ Find node with placeholder text
  │  ├─ Replace with actual value
  │  └─ Preserve all XML structure
  └─ Save modified unpacked/template/

  [pack_docx.py]
  ├─ Re-package unpacked/template/ as DOCX
  └─ Save to output/filled.docx

Output:
  workspaces/{session_id}/output/
  └── filled.docx                    ← Final deliverable
      ├─ All formatting preserved
      ├─ All styles intact
      ├─ Document structure unchanged
      └─ Only text placeholders replaced
```

---

## Usage Examples

### Example 1: Simple Manual Fill

**User Request**: "Fill the template with name='John Smith' and date='2025-11-14'"

**Agent Workflow**:

```python
# Stage 1: Automation
unpack_template("template.docx")
# → Extracts template

find_placeholders()
# → Finds {{name}}, {{date}}

# Stage 2: Analysis
# Agent creates mapping directly from user input

create_replacement_mapping({
    "name": "John Smith",
    "date": "2025-11-14"
})
# → Saves mapping

generate_preview()
# → Shows:
#   {{name}} → John Smith
#   {{date}} → 2025-11-14

# Get user approval
# → User: "Looks good"

# Stage 3: Automation
fill_template()
# → Fills template with values

pack_template()
# → Creates final DOCX
```

**Result**: `workspaces/{session_id}/output/filled.docx`

---

### Example 2: Extract from Source Document

**User Request**: "Fill template.docx using data from invoice.docx"

**Agent Workflow**:

```python
# Stage 1: Automation
unpack_template("template.docx")
# → Extracts template

find_placeholders()
# → Finds {{vendor_name}}, {{invoice_number}}, {{amount}}, {{date}}

convert_docx_to_markdown("invoice.docx", "invoice.md")
# → Converts source document for reading

# Stage 2: Analysis
invoice_content = read_text_file("invoice.md")
# → Reads markdown:
#   # Invoice #12345
#   Vendor: Acme Corp
#   Amount: $50,000
#   Date: 2025-11-14
#   ...

# Agent analyzes and creates mapping:
# "vendor_name" → "Acme Corp" (from "Vendor:" line)
# "invoice_number" → "12345" (from title)
# "amount" → "$50,000" (from "Amount:" line)
# "date" → "2025-11-14" (from "Date:" line)

create_replacement_mapping({
    "vendor_name": "Acme Corp",
    "invoice_number": "12345",
    "amount": "$50,000",
    "date": "2025-11-14"
})

generate_preview()
# → Shows mapping

# Stage 3: Automation (after approval)
fill_template()
pack_template()
```

**Result**: Filled document with vendor details extracted from source

---

### Example 3: Multiple Source Documents

**User Request**: "Use data from company.docx and employee.docx to fill template.docx"

**Agent Workflow**:

```python
# Stage 1: Automation
unpack_template("template.docx")
find_placeholders()
# → Finds {{company}}, {{ceo}}, {{employee_name}}, {{title}}, {{start_date}}

# Convert both sources
convert_docx_to_markdown("company.docx", "company.md")
convert_docx_to_markdown("employee.docx", "employee.md")

# Stage 2: Analysis
company_info = read_text_file("company.md")
employee_info = read_text_file("employee.md")

# Agent combines data from both sources:
# company_info → "{{company}}" and "{{ceo}}"
# employee_info → "{{employee_name}}", "{{title}}", "{{start_date}}"

create_replacement_mapping({
    "company": "Acme Corporation",
    "ceo": "Jane Doe",
    "employee_name": "John Smith",
    "title": "Software Engineer",
    "start_date": "2025-06-01"
})

generate_preview()

# Stage 3: Automation (after approval)
fill_template()
pack_template()
```

**Result**: Document with data from multiple sources intelligently combined

---

## Session Management

### Session Workspace Structure

Each session gets completely isolated workspace:

```
workspaces/
└── {session_id}/
    ├── input/
    │   ├── template.docx           (User uploads)
    │   ├── source.docx             (User uploads)
    │   └── ...other files...
    │
    ├── unpacked/
    │   ├── template/               (Extracted from template.docx)
    │   │   ├── word/
    │   │   │   ├── document.xml    (← Modified by fill_docx.py)
    │   │   │   ├── styles.xml
    │   │   │   └── relationships/
    │   │   ├── _rels/
    │   │   └── [Content_Types].xml
    │   │
    │   └── source/                 (Optional, extracted from source.docx)
    │       └── word/
    │
    ├── debug/
    │   ├── source.md               (Markdown conversion of source)
    │   ├── placeholders.json        (Found placeholders from Stage 1)
    │   ├── replacements.json        (Mapping created in Stage 2)
    │   └── preview.txt             (Generated preview)
    │
    └── output/
        ├── filled.docx             (Final deliverable from Stage 3)
        └── ...other outputs...
```

### Session Isolation

**Why Important**:
- Multiple users can use agent simultaneously
- Each user's data completely separated
- No file interference between sessions
- Clean workspace per session

**How It Works**:
- Each tool call includes `session_id`
- SessionAwareDocxTools maps relative filenames to session-specific paths
- All file operations scoped to: `workspaces/{session_id}/`
- Sessions are independent and concurrent-safe

**Example**:
```python
# User 1 calls: unpack_template("template.docx")
# → Actually processes: workspaces/user1_session_123/input/template.docx
# → Creates: workspaces/user1_session_123/unpacked/template/

# User 2 calls: unpack_template("template.docx")
# → Actually processes: workspaces/user2_session_456/input/template.docx
# → Creates: workspaces/user2_session_456/unpacked/template/

# No interference between users!
```

### Cleanup

User can call `cleanup()` to delete entire session workspace:

```python
cleanup()
# → Deletes: workspaces/{session_id}/
# → Frees disk space
# → Useful after session complete
```

---

## Tool Reference

### Tools by Stage

#### Stage 1: Automation Tools

| Tool | Purpose | Input | Output |
|------|---------|-------|--------|
| `unpack_template(filename)` | Extract DOCX to XML structure | DOCX filename | "✅ Successfully unpacked..." |
| `convert_docx_to_markdown(filename, output_filename)` | Convert DOCX to readable markdown | DOCX filename, output name | "✅ Successfully converted..." |

#### Stage 2: Analysis Tools

| Tool | Purpose | Input | Output |
|------|---------|-------|--------|
| `read_text_file(filename)` | Read markdown/text files for analysis | Filename | File contents (first 10K chars) |
| `read_json_file(filename)` | Read JSON configuration files | Filename | Parsed JSON content |

#### Stage 3: Automation Tools

| Tool | Purpose | Input | Output |
|------|---------|-------|--------|
| `fill_fields(field_mapping)` | Fill empty fields using Document Library | Dict of {label: value} pairs | "✅ Filled X field(s): list... \| Skipped Y: list..." |
| `pack_template(filename)` | Create final DOCX from unpacked XML | Output filename | "✅ Successfully packed to filename.docx" |

#### Utility Tools

| Tool | Purpose | Input | Output |
|------|---------|-------|--------|
| `list_input_files()` | See uploaded files | (none) | List of files in input/ |
| `list_output_files()` | See generated files | (none) | List of files in output/ |
| `get_session_info()` | Get workspace paths and status | (none) | Session details and file counts |
| `cleanup()` | Delete entire session workspace | (none) | "✅ Deleted..." or appropriate message |

---

## Data Structures

### Field Mapping (Stage 2 Output)

The agent creates this dictionary based on semantic analysis:

```python
{
  "Project Manager": "John Smith",
  "Date": "11/14/2025",
  "Budget": "$50,000",
  "Company": "Acme Corporation"
}
```

**Key Points**:
- Keys are label text found in template (e.g., "Project Manager:")
- Values are extracted and formatted values from source
- Ready to pass directly to `fill_fields()` tool
- Agent asks for user confirmation before using

### Fill Fields Tool Response (Stage 3)

After calling `fill_fields(field_mapping)`:

```
✅ Filled 3 field(s): Project Manager, Date, Budget | Skipped 1: Salary
```

**Interpretation**:
- Filled fields: Successfully filled with values
- Skipped fields: No empty field found after label search
- Tool handles formatting preservation automatically

### Session Info (Utility)

```
📋 Session Information:
  - Session ID: 785cc09b-9ade-4aef-b676-b05ee302750d
  - Input directory: workspaces/785cc09b.../input
  - Output directory: workspaces/785cc09b.../output
  - Debug directory: workspaces/785cc09b.../debug
  - Input files: 2 file(s)
  - Output files: 1 file(s)
```

---

## Troubleshooting

### Issue: "pandoc: command not found"

**Cause**: Pandoc not installed on system

**Solution**:
```bash
# Linux
sudo apt-get install pandoc

# macOS
brew install pandoc

# Windows
choco install pandoc
```

### Issue: "File not found: template.docx"

**Cause**: File not uploaded to session

**Solution**:
```python
# Check what files are available
list_input_files()

# Upload missing file via REST API or re-upload
```

### Issue: "No unpacked templates found"

**Cause**: `find_placeholders()` called before `unpack_template()`

**Solution**:
```python
# Always unpack first
unpack_template("template.docx")

# Then find placeholders
find_placeholders()
```

### Issue: Placeholder not being replaced

**Possible Causes**:
1. Placeholder name doesn't match exactly (case-sensitive)
2. Placeholder syntax wrong (must be `{{name}}` not `{name}` or `{{ name }}`)
3. Placeholder not in document.xml (might be in header/footer)

**Solution**:
```python
# Verify found placeholders
placeholders = read_json_file("placeholders.json")
print(placeholders)

# Check replacement mapping has correct names
mapping = read_json_file("replacements.json")
print(mapping)
```

### Issue: Formatting lost after filling

**This should not happen!** The agent uses proper XML replacement that preserves formatting.

**If it occurs**:
1. Check fill_docx.py is using `replace_node()` (it should be)
2. Verify XML structure is valid before filling
3. Try filling with simpler text first

### Issue: Agent doesn't understand what data to extract

**Cause**: Source document structure not clear to agent

**Solution**:
```python
# Provide clearer instructions:
# - Use consistent headings
# - Label data clearly
# - Or provide manual mapping

# Alternative: Create mapping manually
mapping = {
    "name": "John Smith",
    "company": "Acme Corp"
}
create_replacement_mapping(mapping)
generate_preview()
fill_template()
```

### Issue: Session takes too long to cleanup

**Cause**: Large documents create large unpacked directories

**Solution**:
- Sessions automatically cleanup after inactivity
- Manual cleanup available: `cleanup()`
- Large files (>50MB) handled correctly but take time

---

## Advanced Usage

### Custom Placeholder Formats

Placeholders must match pattern: `{{alphanumeric_underscore}}`

Valid examples:
- `{{name}}`
- `{{first_name}}`
- `{{company_name_full}}`
- `{{date_2025}}`

Invalid examples:
- `{{first-name}}` (dashes not allowed)
- `{{first name}}` (spaces not allowed)
- `{{ name }}` (spaces around not allowed)

### Handling Missing Data

If placeholder not found in source:

```python
# Option 1: Provide manual value
mapping = {
    "name": "Unknown",
    "company": "Not provided"
}

# Option 2: Ask user
# "I couldn't find 'company' in source. What should I use?"

# Option 3: Skip placeholder
# Leave it as {{placeholder}} in document
```

### Batch Processing Multiple Templates

```python
# User: "Fill invoice_template.docx and statement_template.docx using company.docx"

# Stage 1: Unpack both
unpack_template("invoice_template.docx")
unpack_template("statement_template.docx")

# Both extract and analyze source once
convert_docx_to_markdown("company.docx", "company.md")

# Stage 2: Create separate mappings for each
# (Different placeholders in each template)

# Stage 3: Fill and pack both
fill_template()  # Uses current working template
pack_template("filled_invoice.docx")

pack_template("filled_statement.docx")
```

---

## Performance Considerations

### File Sizes

- **Input documents**: Up to 50MB supported
- **Unpacked XML**: ~10x uncompressed size (100KB DOCX → 1MB unpacked)
- **Disk space needed**: ~20x input size during processing
  - 10MB DOCX → ~200MB workspace

### Processing Times

| Operation | Time |
|-----------|------|
| Unpack DOCX | 1-5 seconds |
| Extract placeholders | 0.5-2 seconds |
| Convert to markdown | 2-10 seconds (depends on document size) |
| Fill template (100 placeholders) | 1-5 seconds |
| Pack DOCX | 2-8 seconds |
| **Total workflow** | 10-40 seconds |

### Optimization Tips

1. **Batch operations**: Process multiple documents in one session
2. **Reuse conversions**: If filling multiple templates with same source, only convert once
3. **Cleanup sessions**: Delete old sessions to free disk space
4. **Monitor disk**: Large concurrent sessions use significant disk

---

## Summary

The DOCX Autofill Agent follows a proven three-stage workflow:

1. **Stage 1 (AUTOMATION)**: Unpack documents, extract placeholders, convert sources to readable format
2. **Stage 2 (LLM ANALYSIS)**: Intelligently analyze content and create placeholder→value mappings
3. **Stage 3 (AUTOMATION)**: Fill templates using mappings and pack to final DOCX

The system provides:
- ✅ Complete session isolation
- ✅ Formatting preservation
- ✅ Intelligent data extraction
- ✅ User-friendly previews
- ✅ Error handling and recovery

See [DEPENDENCIES.md](DEPENDENCIES.md) for installation requirements and [README.md](README.md) for API documentation.
