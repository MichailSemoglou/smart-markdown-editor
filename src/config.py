"""Configuration constants for the Smart Markdown Editor.

This module contains all configurable constants, timing values, and settings
used throughout the application. Centralizing these values makes the codebase
more maintainable and allows for easy customization.
"""

from typing import Final

# =============================================================================
# Application Metadata
# =============================================================================

APP_NAME: Final[str] = "Smart Markdown Editor"
APP_VERSION: Final[str] = "1.0.0"
APP_ORGANIZATION: Final[str] = "smart-markdown-editor"

# =============================================================================
# Timing Constants (in milliseconds)
# =============================================================================

# Delay before updating the HTML preview after text changes
PREVIEW_UPDATE_DELAY_MS: Final[int] = 300

# Delay before updating document analysis after text changes
ANALYSIS_UPDATE_DELAY_MS: Final[int] = 800

# Default auto-save interval
DEFAULT_AUTOSAVE_INTERVAL_MS: Final[int] = 30_000  # 30 seconds

# =============================================================================
# File Size Limits
# =============================================================================

# Maximum file size in bytes (10 MB)
MAX_FILE_SIZE_BYTES: Final[int] = 10 * 1024 * 1024

# Maximum recent files to store
MAX_RECENT_FILES: Final[int] = 10

# =============================================================================
# Editor Settings
# =============================================================================

# Default font settings
DEFAULT_FONT_FAMILY: Final[str] = "SF Mono"
DEFAULT_FONT_SIZE: Final[int] = 14

# Fallback font families
FONT_FALLBACKS: Final[list[str]] = [
    "Monaco", "Menlo", "Consolas", "Courier New", "monospace"
]

# =============================================================================
# Reading/Analysis Constants
# =============================================================================

# Average words per minute for reading time estimation
WORDS_PER_MINUTE: Final[int] = 200

# Maximum recommended paragraph length (words)
MAX_PARAGRAPH_LENGTH_WARNING: Final[int] = 100
MAX_PARAGRAPH_LENGTH_ERROR: Final[int] = 150

# Maximum recommended line length
MAX_LINE_LENGTH_WARNING: Final[int] = 120

# =============================================================================
# Theme Colors
# =============================================================================

class LightTheme:
    """Color constants for light theme."""
    EDITOR_BG: Final[str] = "#ffffff"
    EDITOR_FG: Final[str] = "#333333"
    EDITOR_BORDER: Final[str] = "#ddd"
    SELECTION_BG: Final[str] = "#0078d4"
    SELECTION_FG: Final[str] = "#ffffff"

    # Syntax highlighting colors
    HEADING_COLOR: Final[str] = "#0b4f9c"
    MUTED_COLOR: Final[str] = "#6a737d"
    RULE_COLOR: Final[str] = "#d0d7de"
    CODE_FG: Final[str] = "#0969da"
    CODE_BG: Final[str] = "#f6f8fa"
    LINK_COLOR: Final[str] = "#0969da"
    URL_COLOR: Final[str] = "#1a7f37"


class DarkTheme:
    """Color constants for dark theme."""
    EDITOR_BG: Final[str] = "#0d1117"
    EDITOR_FG: Final[str] = "#c9d1d9"
    EDITOR_BORDER: Final[str] = "#30363d"
    SELECTION_BG: Final[str] = "#2f81f7"
    SELECTION_FG: Final[str] = "#ffffff"

    # Syntax highlighting colors
    HEADING_COLOR: Final[str] = "#2f81f7"
    MUTED_COLOR: Final[str] = "#8b949e"
    RULE_COLOR: Final[str] = "#30363d"
    CODE_FG: Final[str] = "#c9d1d9"
    CODE_BG: Final[str] = "#161b22"
    LINK_COLOR: Final[str] = "#2f81f7"
    URL_COLOR: Final[str] = "#3fb950"


# =============================================================================
# Preview CSS Constants
# =============================================================================

# Maximum content width in preview
PREVIEW_MAX_WIDTH: Final[str] = "800px"
PREVIEW_PADDING: Final[str] = "20px"

# =============================================================================
# Export Settings
# =============================================================================

# Supported export formats
SUPPORTED_EXPORT_FORMATS: Final[list[str]] = [
    "md", "txt", "html", "docx", "pdf", "rtf", "odt"
]

# File filter patterns for export dialogs
EXPORT_FILE_FILTERS: Final[dict[str, str]] = {
    "md": "Markdown Files (*.md)",
    "txt": "Text Files (*.txt)",
    "html": "HTML Files (*.html)",
    "docx": "Word Documents (*.docx)",
    "pdf": "PDF Files (*.pdf)",
    "rtf": "Rich Text Format (*.rtf)",
    "odt": "OpenDocument Text (*.odt)",
}

# =============================================================================
# Quality Thresholds
# =============================================================================

# Readability score thresholds
READABILITY_EXCELLENT: Final[int] = 80
READABILITY_GOOD: Final[int] = 60
READABILITY_POOR: Final[int] = 40

# =============================================================================
# Window Settings
# =============================================================================

# Default window geometry
DEFAULT_WINDOW_WIDTH: Final[int] = 1400
DEFAULT_WINDOW_HEIGHT: Final[int] = 900

# Default splitter sizes (editor, preview, assistant)
DEFAULT_SPLITTER_SIZES: Final[list[int]] = [420, 630, 350]

# =============================================================================
# Settings Keys
# =============================================================================

class SettingsKeys:
    """Keys used for QSettings storage."""
    DARK_MODE: Final[str] = "darkMode"
    PREVIEW_CSS_PATH: Final[str] = "previewCssPath"
    AUTOSAVE_INTERVAL_MS: Final[str] = "autoSaveIntervalMs"
    RECENT_FILES: Final[str] = "recentFiles"
