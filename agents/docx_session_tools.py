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

    def extract_all_data(self, source_filename: str) -> Dict[str, Any]:
        """
        Phase 2: Extract data from source document using 3 methods.

        Implements AUTO_FILL_WORKFLOW Phase 2:
        - Text content extraction
        - Table structure extraction
        - Structured Data Tag (SDT) form field extraction

        Args:
            source_filename: Filename in input directory

        Returns:
            Dict with extracted data: {text, tables, sdt_fields, extracted_values}
        """
        input_dir = self.workspace.get_input_dir(self.session_id)
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)
        debug_dir = self.workspace.get_debug_dir(self.session_id)

        source_path = input_dir / source_filename
        if not source_path.exists():
            return {
                'error': f'File not found: {source_filename}',
                'text': '',
                'tables': [],
                'sdt_fields': {},
                'extracted_values': {}
            }

        # Unpack source if needed
        source_unpacked = unpacked_dir / 'source'
        if not source_unpacked.exists():
            docx_tools.unpack_docx(str(source_path), str(source_unpacked))

        # Extract using all methods and save results to debug directory
        data = docx_tools.extract_all_data(str(source_unpacked), str(debug_dir))

        self.workspace.update_last_accessed(self.session_id)
        return data

    def fill_fields(self, field_mapping: dict) -> Dict[str, Any]:
        """
        Phase 3 & 5: Fill DOCX fields using multi-strategy approach with validation.

        Implements AUTO_FILL_WORKFLOW Phase 3 filling with 5 implemented strategies:
        - Strategy A: Text placeholder replacement ({{FIELD}})
        - Strategy B: Structured Data Tag (SDT) replacement
        - Strategy C: Element ID-based targeting
        - Strategy D: Multi-run placeholder handling
        - Strategy E: Table row filling
        - Strategy F: Conditional content (NOT YET IMPLEMENTED - stub only)

        Automatically validates with Phase 5 validation (3-tier).

        Args:
            field_mapping: Dict mapping placeholder names to values

        Returns:
            Dict with fill results: {status, filled, skipped, strategies_used, validation_passed, summary}
        """
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)
        debug_dir = self.workspace.get_debug_dir(self.session_id)

        # Find unpacked subdirectory
        subdirs = [d for d in unpacked_dir.iterdir() if d.is_dir()]
        if not subdirs:
            return {
                'status': 'failed',
                'error': 'No unpacked templates found',
                'summary': '❌ No unpacked templates found'
            }

        result = docx_tools.fill_fields(str(subdirs[0]), field_mapping, str(debug_dir))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def insert_placeholders(self, field_analysis: Dict[str, Any]) -> Dict[str, Any]:
        """
        Phase 1.3: Insert {{FIELD_NAME}} placeholders into template.

        Automatically detects where fields should go and inserts {{FIELD_NAME}} placeholders.

        Args:
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
            Dict with insertion results: {status, inserted_count, inserted_fields, summary}
        """
        unpacked_dir = self.workspace.get_unpacked_dir(self.session_id)
        debug_dir = self.workspace.get_debug_dir(self.session_id)

        # Find unpacked subdirectory
        subdirs = [d for d in unpacked_dir.iterdir() if d.is_dir()]
        if not subdirs:
            return {
                'status': 'failed',
                'error': 'No unpacked templates found',
                'summary': '❌ No unpacked templates found'
            }

        result = docx_tools.insert_placeholders(str(subdirs[0]), field_analysis, str(debug_dir))
        self.workspace.update_last_accessed(self.session_id)
        return result

    def cleanup(self) -> str:
        """Clean up session workspace."""
        if self.workspace.cleanup_session(self.session_id):
            return f"✅ Cleaned up session {self.session_id}"
        else:
            return f"❌ Failed to cleanup session {self.session_id}"

    def read_lib_file(self, filename: str) -> str:
        """Read a file from the project's lib/ directory.

        Args:
            filename: Name of file in lib/ (e.g., "docx-js.md")

        Returns:
            File contents or error message
        """
        import os
        lib_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'lib', filename)

        try:
            if not os.path.exists(lib_path):
                return f"❌ File not found: lib/{filename}"

            with open(lib_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return content
        except Exception as e:
            return f"❌ Error reading lib/{filename}: {e}"

    def save_debug_file(self, filename: str, content: str) -> str:
        """Save a file to the session's debug directory.

        Args:
            filename: Name to save as (e.g., "generate_docx.js")
            content: File content to save

        Returns:
            Path where file was saved or error message
        """
        import os
        debug_dir = self.workspace.get_debug_dir(self.session_id)

        try:
            os.makedirs(debug_dir, exist_ok=True)
            filepath = os.path.join(debug_dir, filename)

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)

            self.workspace.update_last_accessed(self.session_id)
            return f"✅ Saved: {filepath}"
        except Exception as e:
            return f"❌ Error saving {filename}: {e}"
