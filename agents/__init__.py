"""DOCX Autofill Agent module"""
from .session_workspace import SessionWorkspaceManager
from .docx_session_tools import SessionAwareDocxTools
from .docx_agent_with_sessions import docx_agent_with_sessions

__all__ = [
    "SessionWorkspaceManager",
    "SessionAwareDocxTools",
    "docx_agent_with_sessions",
]
