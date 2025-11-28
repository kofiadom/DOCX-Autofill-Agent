"""
DOCX Session-Aware Tool Functions - Layer 2
Tool functions that receive RunContext and delegate to SessionAwareDocxTools

Enhanced with Phase 2-5 capabilities:
- Phase 2: Data extraction (extract_all_data)
- Phase 3: Multi-strategy filling (fill_fields with 6 strategies)
- Phase 5: 3-tier validation (automatic)
"""
import json
from typing import Dict, Any
from agno.run import RunContext
from .docx_session_tools import SessionAwareDocxTools


def list_input_files(run_context: RunContext) -> str:
    """List all files in the session's input directory."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.list_input_files()


def list_output_files(run_context: RunContext) -> str:
    """List all files in the session's output directory."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.list_output_files()


def get_session_info(run_context: RunContext) -> str:
    """Get session workspace information."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.get_session_info()


def unpack_template(
    run_context: RunContext,
    filename: str
) -> str:
    """Phase 1: Unpack a DOCX template to XML for editing."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.unpack_template(filename)


def convert_docx_to_markdown(
    run_context: RunContext,
    filename: str,
    output_filename: str = None
) -> str:
    """AUTOMATION: Convert source DOCX to markdown for agent analysis."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.convert_docx_to_markdown(filename, output_filename)


def pack_template(
    run_context: RunContext,
    output_filename: str = None
) -> str:
    """Phase 6: Pack filled XML back to DOCX format."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.pack_template(output_filename)


def read_json_file(
    run_context: RunContext,
    filename: str
) -> str:
    """Read a JSON file from the session's debug directory."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.read_json_file(filename)


def read_text_file(
    run_context: RunContext,
    filename: str
) -> str:
    """Read a text file from the session's debug directory."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.read_text_file(filename)


def extract_all_data(run_context: RunContext, source_filename: str) -> str:
    """Phase 2: Extract data from source document using 3 methods.

    Implements AUTO_FILL_WORKFLOW Phase 2 data extraction:
    - Text content extraction (DOM parsing)
    - Table structure extraction (rows/cells)
    - Structured Data Tag (SDT) form field extraction

    Returns:
        JSON string with combined extracted data
    """
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    data = tools.extract_all_data(source_filename)

    # Return as JSON string for agent processing
    return json.dumps(data, indent=2, default=str)


def fill_fields(run_context: RunContext, field_mapping: dict) -> str:
    """Phase 3 & 5: Fill DOCX fields using multi-strategy approach with validation.

    Applies AUTO_FILL_WORKFLOW Phase 3 filling strategies:
    - Strategy A: Text placeholder replacement ({{FIELD}})
    - Strategy B: Structured Data Tag (SDT) replacement
    - Strategy C: Element ID-based targeting
    - Strategy D: Multi-run placeholder handling
    - Strategy E: Table row filling
    - Strategy F: Conditional sections (NOT YET IMPLEMENTED - stub only)

    Automatically validates with Phase 5 validation (3-tier):
    - Tier 1: Placeholder completion
    - Tier 2: Document integrity
    - Tier 3: XML well-formedness

    Returns:
        JSON string with fill results and validation status
    """
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    result = tools.fill_fields(field_mapping)

    # Result is always a dict from new implementation, convert to JSON string for agent
    return json.dumps(result, indent=2, default=str)


def insert_placeholders(run_context: RunContext, field_analysis: dict) -> str:
    """Phase 1.3: Insert {{FIELD_NAME}} placeholders into template.

    Analyzes template structure and inserts placeholders where missing.

    Args:
        run_context: Agent execution context
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

    Returns:
        JSON string with insertion results
    """
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    result = tools.insert_placeholders(field_analysis)
    return json.dumps(result, indent=2, default=str)


def cleanup(run_context: RunContext) -> str:
    """Clean up the session workspace."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.cleanup()


def read_lib_file(run_context: RunContext, filename: str) -> str:
    """Read a file from the lib/ directory (documentation, references, etc.)

    Use this to read:
    - /lib/docx-js.md (docx library reference)
    - /lib/COMPLEXITY_BASED_DECISION_LOGIC.md (decision guide)
    - Other library documentation files

    Args:
        filename: Name of file in lib/ directory (e.g., "docx-js.md")

    Returns:
        File contents as string
    """
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.read_lib_file(filename)


def save_debug_file(run_context: RunContext, filename: str, content: str) -> str:
    """Save a file to the session's debug directory.

    Use this to save JavaScript files before execution:
    - generate_docx.js (for docx-js fallback approach)
    - Other temporary files needed for processing

    Args:
        filename: Name to save as (e.g., "generate_docx.js")
        content: File content to save

    Returns:
        Path where file was saved
    """
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.save_debug_file(filename, content)
