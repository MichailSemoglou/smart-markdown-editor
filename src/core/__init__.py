"""Core modules for the Smart Markdown Editor.

This package contains the core functionality including:
- MarkdownAnalyzer: Document analysis and quality metrics
- MarkdownSyntaxHighlighter: Syntax highlighting for the editor (UI-only, import directly)
"""

from src.core.analyzer import MarkdownAnalyzer

__all__ = ["MarkdownAnalyzer"]
