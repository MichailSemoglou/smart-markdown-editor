"""Markdown syntax highlighter for the text editor.

This module provides syntax highlighting for markdown documents in the
QTextEdit widget, supporting both light and dark themes.

Example:
    >>> from PySide6.QtWidgets import QTextEdit
    >>> from src.core.highlighter import MarkdownSyntaxHighlighter
    >>> editor = QTextEdit()
    >>> highlighter = MarkdownSyntaxHighlighter(editor.document(), dark_mode=False)
"""

from __future__ import annotations

import logging
from typing import List, Tuple

from PySide6.QtGui import (
    QSyntaxHighlighter,
    QTextCharFormat,
    QColor,
    QFont,
)
from PySide6.QtCore import QRegularExpression

from src.config import DarkTheme, LightTheme

# Configure module logger
logger = logging.getLogger(__name__)


class MarkdownSyntaxHighlighter(QSyntaxHighlighter):
    """Basic Markdown syntax highlighting for the editor.
    
    Provides syntax highlighting for common markdown elements including:
    - Headings (H1-H6)
    - Blockquotes
    - List markers (ordered and unordered)
    - Horizontal rules
    - Bold and italic text
    - Inline code
    - Links and URLs
    - Fenced code blocks
    
    Attributes:
        _dark_mode: Whether dark mode is enabled.
        _rule_formats: List of (regex, format) tuples for highlighting.
        _fence_re: Regex for detecting code fence markers.
        _codeblock_format: Format for code block content.
    """
    
    def __init__(self, document, dark_mode: bool = False) -> None:
        """Initialize the syntax highlighter.
        
        Args:
            document: The QTextDocument to highlight.
            dark_mode: Whether to use dark mode colors. Defaults to False.
        """
        super().__init__(document)
        self._dark_mode = bool(dark_mode)
        self._rule_formats: List[Tuple[QRegularExpression, QTextCharFormat]] = []
        self._fence_re = QRegularExpression(r"^\s{0,3}(```|~~~)")
        self._codeblock_format = QTextCharFormat()
        self._build_formats()
        logger.debug(f"MarkdownSyntaxHighlighter initialized (dark_mode={dark_mode})")
    
    def set_dark_mode(self, dark_mode: bool) -> None:
        """Set the dark mode state and re-highlight.
        
        Args:
            dark_mode: Whether to enable dark mode.
        """
        dark_mode = bool(dark_mode)
        if dark_mode == self._dark_mode:
            return
        
        self._dark_mode = dark_mode
        self._build_formats()
        self.rehighlight()
        logger.debug(f"Dark mode changed to {dark_mode}")
    
    def _get_theme_colors(self) -> Tuple[QColor, QColor, QColor, QColor, QColor, QColor, QColor]:
        """Get color values based on current theme.
        
        Returns:
            Tuple of colors: (heading, muted, rule, code_fg, code_bg, link, url)
        """
        if self._dark_mode:
            return (
                QColor(DarkTheme.HEADING_COLOR),
                QColor(DarkTheme.MUTED_COLOR),
                QColor(DarkTheme.RULE_COLOR),
                QColor(DarkTheme.CODE_FG),
                QColor(DarkTheme.CODE_BG),
                QColor(DarkTheme.LINK_COLOR),
                QColor(DarkTheme.URL_COLOR),
            )
        else:
            return (
                QColor(LightTheme.HEADING_COLOR),
                QColor(LightTheme.MUTED_COLOR),
                QColor(LightTheme.RULE_COLOR),
                QColor(LightTheme.CODE_FG),
                QColor(LightTheme.CODE_BG),
                QColor(LightTheme.LINK_COLOR),
                QColor(LightTheme.URL_COLOR),
            )
    
    def _build_formats(self) -> None:
        """Build all highlighting formats based on current theme."""
        self._rule_formats = []
        
        heading_color, muted_color, rule_color, code_fg, code_bg, link_color, url_color = \
            self._get_theme_colors()
        
        # Heading format (H1-H6)
        heading_format = QTextCharFormat()
        heading_format.setForeground(heading_color)
        heading_format.setFontWeight(QFont.Weight.Bold)
        self._add_rule(r"^\s{0,3}#{1,6} .*", heading_format)
        
        # Blockquote format
        blockquote_format = QTextCharFormat()
        blockquote_format.setForeground(muted_color)
        self._add_rule(r"^\s{0,3}>\s.*", blockquote_format)
        
        # List marker format (unordered and ordered)
        list_marker_format = QTextCharFormat()
        list_marker_format.setForeground(muted_color)
        list_marker_format.setFontWeight(QFont.Weight.Bold)
        self._add_rule(r"^\s{0,3}([-*+])\s+", list_marker_format)
        self._add_rule(r"^\s{0,3}(\d+)\.\s+", list_marker_format)
        
        # Horizontal rule format
        hr_format = QTextCharFormat()
        hr_format.setForeground(rule_color)
        self._add_rule(r"^\s{0,3}(-{3,}|\*{3,}|_{3,})\s*$", hr_format)
        
        # Bold format
        bold_format = QTextCharFormat()
        bold_format.setFontWeight(QFont.Weight.Bold)
        self._add_rule(r"\*\*[^\*\n]+\*\*", bold_format)
        self._add_rule(r"__[^_\n]+__", bold_format)
        
        # Italic format
        italic_format = QTextCharFormat()
        italic_format.setFontItalic(True)
        self._add_rule(r"(?<!\*)\*[^\*\n]+\*(?!\*)", italic_format)
        self._add_rule(r"(?<!_)_[^_\n]+_(?!_)", italic_format)
        
        # Inline code format
        inline_code_format = QTextCharFormat()
        inline_code_format.setForeground(code_fg)
        inline_code_format.setBackground(code_bg)
        self._add_rule(r"`[^`\n]+`", inline_code_format)
        
        # Link text format
        link_text_format = QTextCharFormat()
        link_text_format.setForeground(link_color)
        self._add_rule(r"\[[^\]]+\](?=\()", link_text_format)
        
        # Link URL format
        link_url_format = QTextCharFormat()
        link_url_format.setForeground(url_color)
        self._add_rule(r"\([^\)\s]+\)", link_url_format)
        
        # Code block format (for fenced code blocks)
        self._codeblock_format = QTextCharFormat()
        if self._dark_mode:
            self._codeblock_format.setForeground(QColor(DarkTheme.CODE_FG))
            self._codeblock_format.setBackground(QColor(DarkTheme.CODE_BG))
        else:
            self._codeblock_format.setForeground(QColor("#24292f"))
            self._codeblock_format.setBackground(QColor(LightTheme.CODE_BG))
        self._codeblock_format.setFontFamily("SF Mono")
    
    def _add_rule(self, pattern: str, fmt: QTextCharFormat) -> None:
        """Add a highlighting rule.
        
        Args:
            pattern: Regular expression pattern string.
            fmt: Text format to apply to matches.
        """
        self._rule_formats.append((QRegularExpression(pattern), fmt))
    
    def highlightBlock(self, text: str) -> None:
        """Highlight a block of text.
        
        This method is called by Qt for each text block in the document.
        It handles code block state tracking and applies formatting rules.
        
        Args:
            text: The text block to highlight.
        """
        # Check if we're continuing a code block from previous block
        in_code_block = self.previousBlockState() == 1
        
        # Check for fence markers
        fence_match = self._fence_re.match(text)
        is_fence_line = fence_match.hasMatch()
        
        # Handle code block content
        if in_code_block:
            self.setFormat(0, len(text), self._codeblock_format)
            if is_fence_line:
                # Closing fence - end code block
                self.setCurrentBlockState(0)
            else:
                # Continue code block
                self.setCurrentBlockState(1)
            return
        
        if is_fence_line:
            # Opening fence - start code block
            self.setFormat(0, len(text), self._codeblock_format)
            self.setCurrentBlockState(1)
            return
        
        # Normal text - apply highlighting rules
        self.setCurrentBlockState(0)
        for regex, fmt in self._rule_formats:
            it = regex.globalMatch(text)
            while it.hasNext():
                match = it.next()
                start = match.capturedStart()
                length = match.capturedLength()
                if length > 0:
                    self.setFormat(start, length, fmt)
