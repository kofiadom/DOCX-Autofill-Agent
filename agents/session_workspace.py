"""Session-aware workspace management for DOCX autofill agent"""
import os
import shutil
from pathlib import Path
from typing import Optional

class SessionWorkspaceManager:
    """Manages isolated workspace directories for each session"""

    def __init__(self, base_workspace_dir: str = "workspaces"):
        """Initialize workspace manager"""
        self.base_dir = Path(base_workspace_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def get_session_dir(self, session_id: str) -> Path:
        """Get session directory"""
        session_dir = self.base_dir / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        return session_dir

    def get_input_dir(self, session_id: str) -> Path:
        """Get input directory for uploaded files"""
        input_dir = self.get_session_dir(session_id) / "input"
        input_dir.mkdir(parents=True, exist_ok=True)
        return input_dir

    def get_unpacked_dir(self, session_id: str) -> Path:
        """Get directory for unpacked DOCX files"""
        unpacked_dir = self.get_session_dir(session_id) / "unpacked"
        unpacked_dir.mkdir(parents=True, exist_ok=True)
        return unpacked_dir

    def get_debug_dir(self, session_id: str) -> Path:
        """Get directory for debug/analysis files"""
        debug_dir = self.get_session_dir(session_id) / "debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        return debug_dir

    def get_output_dir(self, session_id: str) -> Path:
        """Get directory for output filled DOCX files"""
        output_dir = self.get_session_dir(session_id) / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir

    def list_input_files(self, session_id: str) -> list:
        """List all input files for session"""
        input_dir = self.get_input_dir(session_id)
        if not input_dir.exists():
            return []
        return [f.name for f in input_dir.iterdir() if f.is_file()]

    def list_output_files(self, session_id: str) -> list:
        """List all output files for session"""
        output_dir = self.get_output_dir(session_id)
        if not output_dir.exists():
            return []
        return [f.name for f in output_dir.iterdir() if f.is_file()]

    def cleanup_session(self, session_id: str) -> bool:
        """Delete all files for a session"""
        session_dir = self.get_session_dir(session_id)
        try:
            shutil.rmtree(session_dir)
            return True
        except Exception as e:
            print(f"Error cleaning up session {session_id}: {e}")
            return False

    def get_session_info(self, session_id: str) -> dict:
        """Get information about session workspace"""
        return {
            "session_id": session_id,
            "session_dir": str(self.get_session_dir(session_id)),
            "input_files": self.list_input_files(session_id),
            "output_files": self.list_output_files(session_id),
            "unpacked_dir": str(self.get_unpacked_dir(session_id)),
            "debug_dir": str(self.get_debug_dir(session_id))
        }

    def update_last_accessed(self, session_id: str) -> None:
        """Update last accessed timestamp for a session"""
        # Could be used for cleanup of old sessions
        session_dir = self.get_session_dir(session_id)
        try:
            session_dir.touch(exist_ok=True)
        except:
            pass  # Ignore errors
