"""Export functionality for the Smart Markdown Editor.

This package provides a modular export system with:
- BaseExporter: Abstract base class for all exporters
- MarkdownExporter: Export to markdown format
- TextExporter: Export to plain text
- HTMLExporter: Export as HTML
- DocxExporter: Export to Word documents
- PDFExporter: Export to PDF
- RTFExporter: Export to RTF
- ODTExporter: Export as ODT

The exporters use a registry pattern for dynamic discovery and registration.
"""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from typing import Optional

from pathlib import Path

# Configure module logger
logger = logging.getLogger(__name__)


class BaseExporter(ABC):
    """Abstract base class for all exporters.
    
    All exporters must inherit from this class and implement the export() method.
    The registry pattern allows for dynamic discovery and registration.
    
    Attributes:
        name: Human-readable name of the exporter.
        extension: File extension (without the dot).
        file_filter: File filter string for dialogs.
        requires_library: Optional library name required for this exporter.
        _available: Whether the required library is available.
    """
    
    name: str
    extension: str
    file_filter: str
    requires_library: Optional[str] = None
    _available: bool = True
    
    def __init__(self) -> None:
        """Initialize the exporter and check for required libraries."""
        if self.requires_library:
            self._available = self._check_library()
            if not self._available:
                logger.warning(
                    f"Exporter '{self.name}' unavailable: {self.requires_library} not installed"
                )
    
    @property
    def is_available(self) -> bool:
        """Check if this exporter is available.
        
        Returns:
            bool: True if the exporter can be used.
        """
        return self._available
    
    @abstractmethod
    def export(self, content: str, output_path: Path) -> None:
        """Export the markdown content to the specified format.
        
        Args:
            content: The markdown content to export.
            output_path: The path to write the output file.
        
        Raises:
            ExportError: If export fails.
        """
        pass
    
    def _check_library(self) -> bool:
        """Check if the required library is available.
        
        Returns:
            bool: True if the library is available.
        """
        try:
            assert self.requires_library is not None
            __import__(self.requires_library)
            return True
        except ImportError:
            return False
    
    def get_install_message(self) -> str:
        """Get a message instructing how to install the required library.
        
        Returns:
            str: Installation message or empty string if no library required.
        """
        if self.requires_library:
            return f"Install with: pip install {self.requires_library}"
        return ""


# Global registry of exporters
_EXPORTERS: dict[str, type[BaseExporter]] = {}


def register_exporter(exporter: type[BaseExporter]) -> None:
    """Register an exporter in the global registry.

    Args:
        exporter: The exporter class to register.
    """
    global _EXPORTERS
    _EXPORTERS[exporter.extension] = exporter
    logger.info(f"Registered exporter: {exporter.name} ({exporter.extension})")


def get_exporter(extension: str) -> Optional[type[BaseExporter]]:
    """Get an exporter by file extension.

    Args:
        extension: The file extension (without the dot).

    Returns:
        Exporter class, or None if not found.
    """
    return _EXPORTERS.get(extension)


def get_available_exporters() -> list[type[BaseExporter]]:
    """Get all registered exporters whose required library is present.

    Returns:
        list: List of available exporter classes.
    """
    available = []
    for exporter_class in _EXPORTERS.values():
        try:
            if exporter_class().is_available:
                available.append(exporter_class)
        except Exception:
            pass
    return available


def get_all_exporters() -> list[type[BaseExporter]]:
    """Get all registered exporters (including unavailable ones).

    Returns:
        list: List of all exporter classes.
    """
    return list(_EXPORTERS.values())
