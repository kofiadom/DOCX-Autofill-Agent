"""
DOCX Session-Aware Tool Functions - Layer 2
Tool functions that receive RunContext and delegate to SessionAwareDocxTools
"""
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
    """AUTOMATION STAGE 1: Unpack a DOCX template to XML."""
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
    """AUTOMATION STAGE 3: Pack filled template back to DOCX."""
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


def fill_fields(run_context: RunContext, field_mapping: dict) -> str:
    """STAGE 3: Fill DOCX fields using Document Library and label-based mapping."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.fill_fields(field_mapping)


def cleanup(run_context: RunContext) -> str:
    """Clean up the session workspace."""
    tools = SessionAwareDocxTools(
        session_id=run_context.session_id,
        user_id=run_context.user_id
    )
    return tools.cleanup()
