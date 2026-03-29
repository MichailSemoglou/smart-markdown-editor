"""Utility functions for the Smart Markdown Editor.

This package provides common utility functions including:
- File operations (read, write, validate)
- Input validation and sanitization
- Security helpers

Example:
    >>> from src.utils import validate_file_path, sanitize_filename
    >>> path = validate_file_path("/path/to/file.md")
    >>> safe_name = sanitize_filename("../../../etc/passwd")
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Union

# Configure module logger
logger = logging.getLogger(__name__)

# Constants
MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024  # 10MB
ALLOWED_EXTENSIONS = {'.md', '.txt', '.markdown'}


class FileSizeError(Exception):
    """Raised when a file exceeds the maximum size limit."""
    pass


class SecurityError(Exception):
    """Raised when a security violation is detected."""
    pass


class FileAccessError(Exception):
    """Raised when file access fails."""
    pass


class ValidationError(Exception):
    """Raised when input validation fails."""
    pass


def validate_file_path(file_path: Union[str, Path]) -> Path:
    """Validate a file path for security and accessibility.
    
    Args:
        file_path: Path to validate.
        
    Returns:
        Validated Path object.
        
    Raises:
        FileAccessError: If the file doesn't exist or isn't accessible.
        ValidationError: If the file extension is not allowed.
    """
    path = Path(file_path)
    
    # Resolve to absolute path
    try:
        path = path.resolve()
    except OSError as e:
        logger.error(f"Failed to resolve path: {e}")
        raise FileAccessError(f"Cannot resolve path: {file_path}")
    
    # Check if path exists
    if not path.exists():
        logger.warning(f"File does not exist: {file_path}")
        raise FileAccessError(f"File not found: {file_path}")
    
    # Check if it's a file (not directory)
    if not path.is_file():
        logger.warning(f"Path is not a file: {file_path}")
        raise FileAccessError(f"Path is not a file: {file_path}")
    
    # Validate file extension
    ext = path.suffix.lower()
    if ext and ext not in ALLOWED_EXTENSIONS:
        logger.warning(f"Potentially unsupported file extension: {ext}")
        # Don't raise error, just warn - user may want to open any text file
    
    # Check file is readable
    if not os.access(path, os.R_OK):
        logger.warning(f"File is not readable: {file_path}")
        raise FileAccessError(f"File is not readable: {file_path}")
    
    logger.debug(f"File path validated: {path}")
    return path


def validate_file_size(file_path: Union[str, Path], max_size: int = MAX_FILE_SIZE_BYTES) -> None:
    """Validate file size is within limits.
    
    Args:
        file_path: Path to the file.
        max_size: Maximum allowed size in bytes.
        
    Raises:
        FileSizeError: If file exceeds size limit.
        FileAccessError: If file size cannot be determined.
    """
    path = Path(file_path)
    
    try:
        size = path.stat().st_size
    except OSError as e:
        logger.error(f"Cannot determine file size: {e}")
        raise FileAccessError(f"Cannot determine file size: {file_path}")
    
    if size > max_size:
        size_mb = size / (1024 * 1024)
        max_mb = max_size / (1024 * 1024)
        logger.warning(f"File too large: {size_mb:.2f}MB > {max_mb:.2f}MB")
        raise FileSizeError(
            f"File exceeds maximum size ({size_mb:.2f}MB > {max_mb:.2f}MB)"
        )
    
    logger.debug(f"File size OK: {size} bytes")


def sanitize_filename(filename: str) -> str:
    """Sanitize a filename to prevent path traversal and invalid characters.
    
    Removes or replaces potentially dangerous characters.
    
    Args:
        filename: The filename to sanitize.
        
    Returns:
        Sanitized filename safe for filesystem use.
    """
    if not filename:
        return "untitled"
    
    # Remove null bytes
    filename = filename.replace('\x00', '')
    
    # Remove path separators and dangerous characters
    dangerous_chars = '<>:"|*?\\/'
    for char in dangerous_chars:
        filename = filename.replace(char, '_')
    
    # Remove control characters (0x00-0x1F)
    filename = re.sub(r'[\x00-\x1f]', '', filename)
    
    # Prevent path traversal - remove .. and leading dots/slashes
    filename = filename.replace('..', '')
    filename = filename.lstrip('./\\')
    
    # Limit length
    max_length = 255
    if len(filename) > max_length:
        name, ext = os.path.splitext(filename)
        filename = name[:max_length - len(ext)] + ext
    
    # Ensure not empty after sanitization
    if not filename or filename.isspace():
        filename = "untitled"
    
    logger.debug(f"Sanitized filename: {filename}")
    return filename


def read_file_content(file_path: Union[str, Path], encoding: str = 'utf-8') -> str:
    """Read file content with validation.
    
    Args:
        file_path: Path to the file.
        encoding: Character encoding (default: utf-8).
        
    Returns:
        File content as string.
        
    Raises:
        FileAccessError: If file cannot be read.
        FileSizeError: If file is too large.
    """
    path = validate_file_path(file_path)
    validate_file_size(path)
    
    try:
        content = path.read_text(encoding=encoding)
        logger.info(f"Read {len(content)} characters from {path}")
        return content
    except UnicodeDecodeError as e:
        logger.error(f"Encoding error reading file: {e}")
        raise FileAccessError(f"Cannot read file with {encoding} encoding: {file_path}")
    except OSError as e:
        logger.error(f"Error reading file: {e}")
        raise FileAccessError(f"Cannot read file: {file_path}")


def write_file_content(
    file_path: Union[str, Path], 
    content: str, 
    encoding: str = 'utf-8'
) -> None:
    """Write content to a file safely.
    
    Args:
        file_path: Path to write to.
        content: Content to write.
        encoding: Character encoding (default: utf-8).
        
    Raises:
        FileAccessError: If file cannot be written.
    """
    path = Path(file_path)
    
    # Ensure parent directory exists
    path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        path.write_text(content, encoding=encoding)
        logger.info(f"Wrote {len(content)} characters to {path}")
    except OSError as e:
        logger.error(f"Error writing file: {e}")
        raise FileAccessError(f"Cannot write file: {file_path}")


def get_unique_filename(directory: Union[str, Path], base_name: str, extension: str) -> Path:
    """Generate a unique filename in a directory.
    
    Args:
        directory: Target directory.
        base_name: Base filename without extension.
        extension: File extension (with or without dot).
        
    Returns:
        Path to a unique file that doesn't exist yet.
    """
    dir_path = Path(directory)
    ext = extension if extension.startswith('.') else f'.{extension}'
    
    # Sanitize base name
    safe_base = sanitize_filename(base_name)
    
    # Try original name first
    candidate = dir_path / f"{safe_base}{ext}"
    if not candidate.exists():
        return candidate
    
    # Add counter suffix
    counter = 1
    while True:
        candidate = dir_path / f"{safe_base}_{counter}{ext}"
        if not candidate.exists():
            return candidate
        counter += 1
        
        # Safety limit
        if counter > 10000:
            raise ValidationError("Cannot generate unique filename")


# Export public API
__all__ = [
    # Exceptions
    'FileSizeError',
    'SecurityError',
    'FileAccessError',
    'ValidationError',
    # Constants
    'MAX_FILE_SIZE_BYTES',
    'ALLOWED_EXTENSIONS',
    # Functions
    'validate_file_path',
    'validate_file_size',
    'sanitize_filename',
    'read_file_content',
    'write_file_content',
    'get_unique_filename',
]
