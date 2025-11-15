"""PowerPoint Generator AgentOS"""

from pathlib import Path
from dotenv import load_dotenv
from fastapi import File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import logging
import re
import uuid

# Load environment variables from .env file
load_dotenv()

from agno.os import AgentOS

# Use the session-aware agent for proper session isolation
from agents.docx_agent_with_sessions import docx_agent_with_sessions
from agents.session_workspace import SessionWorkspaceManager

# Create the AgentOS with session-isolated agent
agent_os = AgentOS(
    os_id="docx-autofill-os",
    agents=[docx_agent_with_sessions],
)
app = agent_os.get_app()

# Configure logging
logger = logging.getLogger(__name__)

# File upload configuration
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_EXTENSIONS = {'.docx', '.doc', '.pptx', '.txt', '.md', '.pdf'}
ALLOWED_MIME_TYPES = {
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'text/plain',
    'text/markdown',
}

# Startup event to load MCP tools
@app.on_event("startup")
async def startup_event():
	print("Loading Agent startup...")

# CORS for frontend
# app.add_middleware(
# 	CORSMiddleware,
# 	allow_origins= ["*"],
# 	allow_credentials=True,
# 	allow_methods=["*"],
# 	allow_headers=["*"],
# )

def sanitize_filename(filename: str) -> str:
	"""
	Sanitize filename to prevent path traversal and other security issues.

	Args:
		filename: Original filename

	Returns:
		Sanitized filename safe for filesystem storage
	"""
	# Remove any path components
	filename = Path(filename).name

	# Remove or replace dangerous characters
	filename = re.sub(r'[^\w\s\-.]', '_', filename)

	# Remove leading/trailing dots and spaces
	filename = filename.strip('. ')

	# Ensure filename is not empty after sanitization
	if not filename:
		filename = f"unnamed_{uuid.uuid4().hex[:8]}"

	return filename


def validate_file(file: UploadFile, content: bytes) -> None:
	"""
	Validate file type and size.

	Args:
		file: Uploaded file object
		content: File content bytes

	Raises:
		HTTPException: If validation fails
	"""
	# Check file size
	if len(content) > MAX_FILE_SIZE:
		raise HTTPException(
			status_code=413,
			detail=f"File '{file.filename}' exceeds maximum size of {MAX_FILE_SIZE // (1024*1024)}MB"
		)

	# Check file extension
	file_ext = Path(file.filename).suffix.lower()
	if file_ext not in ALLOWED_EXTENSIONS:
		raise HTTPException(
			status_code=400,
			detail=f"File type '{file_ext}' not allowed. Supported types: {', '.join(ALLOWED_EXTENSIONS)}"
		)

	# Check MIME type
	if file.content_type and file.content_type not in ALLOWED_MIME_TYPES:
		raise HTTPException(
			status_code=400,
			detail=f"MIME type '{file.content_type}' not allowed"
		)

	# Check for empty files
	if len(content) == 0:
		raise HTTPException(
			status_code=400,
			detail=f"File '{file.filename}' is empty"
		)


def generate_unique_filename(workspace: SessionWorkspaceManager, filename: str) -> str:
	"""
	Generate a unique filename if a file with the same name already exists.

	Args:
		workspace: Session workspace manager
		filename: Desired filename

	Returns:
		Unique filename (may have UUID prefix if duplicate exists)
	"""
	filepath = workspace.get_input_path(filename)

	# If file doesn't exist, return original filename
	if not filepath.exists():
		return filename

	# Generate unique filename with UUID prefix
	file_stem = Path(filename).stem
	file_ext = Path(filename).suffix
	unique_filename = f"{uuid.uuid4().hex[:8]}_{file_stem}{file_ext}"

	logger.info(f"File '{filename}' already exists, renamed to '{unique_filename}'")

	return unique_filename


# Custom file upload endpoint
@app.post("/api/upload/{session_id}")
async def upload_files(
	session_id: str,
	files: List[UploadFile] = File(...)
):
	"""
	Upload files to a session's workspace with validation and security checks.

	Args:
		session_id: Session identifier for workspace isolation
		files: List of files to upload

	Returns:
		JSON response with upload results

	Raises:
		HTTPException: If validation fails or errors occur
	"""
	# Validate session_id format (alphanumeric and hyphens only)
	if not re.match(r'^[\w\-]+$', session_id):
		raise HTTPException(
			status_code=400,
			detail="Invalid session_id format"
		)

	# Check if files are provided
	if not files:
		raise HTTPException(
			status_code=400,
			detail="No files provided"
		)

	# Initialize workspace
	try:
		workspace = SessionWorkspaceManager(session_id=session_id)
	except Exception as e:
		logger.error(f"Failed to initialize workspace for session {session_id}: {e}")
		raise HTTPException(
			status_code=500,
			detail="Failed to initialize session workspace"
		)

	uploaded = []
	failed = []

	# Process each file
	for file in files:
		try:
			# Read file content
			content = await file.read()

			# Validate file
			validate_file(file, content)

			# Sanitize filename
			safe_filename = sanitize_filename(file.filename)

			# Generate unique filename if duplicate exists
			unique_filename = generate_unique_filename(workspace, safe_filename)

			# Save file to workspace
			filepath = workspace.get_input_path(unique_filename)
			with open(filepath, 'wb') as f:
				f.write(content)

			uploaded.append({
				"original_filename": file.filename,
				"stored_filename": unique_filename,
				"size": len(content),
				"content_type": file.content_type,
				"relative_path": f"input/{unique_filename}"
			})

			logger.info(f"Uploaded {file.filename} ({len(content)} bytes) to {filepath}")

		except HTTPException:
			# Re-raise HTTP exceptions (validation errors)
			raise
		except Exception as e:
			# Log and collect other errors
			logger.error(f"Error uploading {file.filename}: {e}")
			failed.append({
				"filename": file.filename,
				"error": str(e)
			})

	# If no files were uploaded successfully, return error
	if not uploaded and failed:
		raise HTTPException(
			status_code=500,
			detail=f"Failed to upload all files: {failed}"
		)

	# Return success response
	response = {
		"success": True,
		"session_id": session_id,
		"uploaded": uploaded,
		"message": f"Successfully uploaded {len(uploaded)} file(s)"
	}

	# Include failed files if any
	if failed:
		response["failed"] = failed
		response["message"] += f", {len(failed)} file(s) failed"

	return response


if __name__ == "__main__":
    # Run the AgentOS server
    agent_os.serve(app="app.main:app", reload=True)
