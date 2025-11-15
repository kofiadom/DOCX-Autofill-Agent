"""
DOCX Session-Aware Tools - Layer 3
SessionAwareDocxTools class that wraps core tools and applies workspace isolation
"""
from pathlib import Path
from typing import Optional, Dict, Any
from .session_workspace import SessionWorkspaceManager
from . import docx_tools  # Import Layer 1 core tools


class SessionAwareDocxTools:
    """
    Wraps DOCX tools and applies session-based workspace isolation.
    Converts relative filenames to absolute session-specific paths.
    """

    def __init__(self, session_id: str, user_id: Optional[str] = None, workspace_base: str = "workspaces"):
        """Initialize session-aware tools."""
        self.session_id = session_id
        self.user_id = user_id
        self.workspace = SessionWorkspaceManager(base_workspace_dir=workspace_base)

    def list_input_files(self) -> str:
        """List all files in session's input directory."""
        files = self.workspace.list_input_files(self.session_id)
        if not files:
            return "📁 No files uploaded to this session yet."

        lines = ["📁 **Files in session input directory:**"]
        for f in sorted(files):
            lines.append(f"  - {f}")

        self.workspace.update_last_accessed(self.session_id)
        return "\n".join(lines)

    def list_output_files(self) -> str:
        """List all files in session's output directory."""
        files = self.workspace.list_output_files(self.session_id)
        if not files:
            return "📁 No output files generated yet."

        lines = ["📁 **Files in session output directory:**"]
        for f in sorted(files):
            lines.append(f"  - {f}")

        self.workspace.update_last_accessed(self.session_id)
        return "\n".join(lines)

    def get_session_info(self) -> str:
        """Get session workspace information."""
        info = self.workspace.get_session_info(self.session_id)
        lines = [
            "📋 **Session Information:**",
            f"  - Session ID: {info['session_id']}",
            f"  - Input directory: {info['session_dir']}/input",
            f"  - Output directory: {info['session_dir']}/output",
            f"  - Debug directory: {info['session_dir']}/debug",
            f"  - Input files: {len(info['input_files'])} file(s)",
            f"  - Output files: {len(info['output_files'])} file(s)"
        ]
        self.workspace.update_last_accessed(self.session_id)
        return "\n".join(lines)

    def unpack_template(self, filename: str) -> str:
        """AUTOMATION: Unpack DOCX template to XML."""
        input_dir = self.workspace.get_input_dir(self.session_id)
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)

        docx_path = input_dir / filename
        if not docx_path.exists():
            return f"❌ File not found: {filename}\n\nUse list_input_files() to see available files."

        output_subdir = unpacked_dir / Path(filename).stem

        result = docx_tools.unpack_docx(str(docx_path), str(output_subdir))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def convert_docx_to_markdown(self, filename: str, output_filename: str = None) -> str:
        """AUTOMATION: Convert source DOCX to markdown for agent analysis."""
        input_dir = self.workspace.get_input_dir(self.session_id)
        debug_dir = self.workspace.get_debug_dir(self.session_id)

        docx_path = input_dir / filename
        if not docx_path.exists():
            return f"❌ File not found: {filename}"

        if not output_filename:
            output_filename = Path(filename).stem + ".md"

        md_path = debug_dir / output_filename

        result = docx_tools.convert_docx_to_markdown(str(docx_path), str(md_path))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def pack_template(self, output_filename: str = None) -> str:
        """AUTOMATION: Pack filled template back to DOCX."""
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)
        output_dir = self.workspace.get_output_dir(self.session_id)

        # Find unpacked directory
        subdirs = [d for d in unpacked_dir.iterdir() if d.is_dir()]
        if not subdirs:
            return "❌ No unpacked templates found."

        target_dir = subdirs[0]

        if not output_filename:
            output_filename = f"{target_dir.name}_filled.docx"

        output_path = output_dir / output_filename

        result = docx_tools.pack_docx(str(target_dir), str(output_path))
        self.workspace.update_last_accessed(self.session_id)

        if "✅" in result:
            return f"{result}\n\n📥 **Download your filled document from:**\n  `workspaces/{self.session_id}/output/{output_filename}`"
        return result

    def read_json_file(self, filename: str) -> str:
        """Read a JSON file from debug directory."""
        debug_dir = self.workspace.get_debug_dir(self.session_id)
        file_path = debug_dir / filename

        result = docx_tools.read_json_file(str(file_path))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def read_text_file(self, filename: str) -> str:
        """Read a text file from debug directory."""
        debug_dir = self.workspace.get_debug_dir(self.session_id)
        file_path = debug_dir / filename

        result = docx_tools.read_text_file(str(file_path))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def fill_fields(self, field_mapping: dict) -> str:
        """STAGE 3: Fill DOCX fields using Document Library."""
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)

        # Find unpacked subdirectory
        subdirs = [d for d in unpacked_dir.iterdir() if d.is_dir()]
        if not subdirs:
            return "❌ No unpacked templates found"

        result = docx_tools.fill_fields(str(subdirs[0]), field_mapping)
        self.workspace.update_last_accessed(self.session_id)
        return result

    def cleanup(self) -> str:
        """Clean up session workspace."""
        if self.workspace.cleanup_session(self.session_id):
            return f"✅ Cleaned up session {self.session_id}"
        else:
            return f"❌ Failed to cleanup session {self.session_id}"
