"""Core modules for the Smart Markdown Editor.

This package contains the core functionality including:
- MarkdownAnalyzer: Document analysis and quality metrics
- MarkdownSyntaxHighlighter: Syntax highlighting for the editor
"""

from src.core.analyzer import MarkdownAnalyzer
from src.core.highlighter import MarkdownSyntaxHighlighter

__all__ = ["MarkdownAnalyzer", "MarkdownSyntaxHighlighter"]
