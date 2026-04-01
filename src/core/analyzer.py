"""Markdown document analyzer for structure and quality metrics.

This module provides the MarkdownAnalyzer class which performs comprehensive
analysis of markdown documents including word count, structure analysis,
readability scoring, and issue detection.

Example:
    >>> from src.core.analyzer import MarkdownAnalyzer
    >>> analyzer = MarkdownAnalyzer("# Hello\\n\\nWorld")
    >>> metrics = analyzer.analyze()
    >>> print(metrics['word_count'])
    2
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from typing import Any

import textstat

from src.config import (
    MAX_LINE_LENGTH_WARNING,
    READABILITY_GREEN,
    READABILITY_ORANGE,
    WORDS_PER_MINUTE,
)

# Configure module logger
logger = logging.getLogger(__name__)


# Pre-compiled regex patterns for performance
# These are compiled once at module load time rather than on each method call
RE_CODE_BLOCK = re.compile(r'```.*?```', re.DOTALL)
RE_INLINE_CODE = re.compile(r'`[^`\n]+`')
RE_WORD = re.compile(r'\b\w+\b')
RE_HEADING = re.compile(r'^(#{1,6})\s+(.+)$')
RE_LINK = re.compile(r'\[([^\]]+)\]\(([^\)]+)\)')
RE_IMAGE = re.compile(r'!\[([^\]]*)\]\(([^\)]+)\)')
RE_FENCED_CODE = re.compile(r'```')
RE_UNORDERED_LIST = re.compile(r'^\s*[-*+]\s+')
RE_ORDERED_LIST = re.compile(r'^\s*\d+\.\s+')
RE_BLOCKQUOTE = re.compile(r'^>\s')
RE_EMPTY_LINK = re.compile(r'\[([^\]]+)\]\(\s*\)')
RE_BOLD = re.compile(r'\*\*[^\*\n]+\*\*|__[^_\n]+__')
RE_ITALIC = re.compile(r'(?<!\*)\*[^\*\n]+\*(?!\*)|(?<!_)_[^_\n]+_(?!_)')


class MarkdownAnalyzer:
    """Analyze markdown text and produce structure/quality metrics.

    This class provides comprehensive analysis of markdown documents including:
    - Basic statistics (word count, character count, line count)
    - Structure analysis (headings, links, images, code blocks, etc.)
    - Quality metrics (readability score, structure quality)
    - Issue detection (empty links, duplicate headings, long lines)

    Attributes:
        text: The markdown text to analyze.
        lines: List of lines in the document.

    Example:
        >>> analyzer = MarkdownAnalyzer("# Title\\n\\nParagraph text.")
        >>> metrics = analyzer.analyze()
        >>> print(metrics['word_count'])
        3
    """

    def __init__(self, text: str) -> None:
        """Initialize the analyzer with markdown text.

        Args:
            text: The markdown text to analyze. None is treated as empty string.
        """
        if text is None:
            logger.warning("MarkdownAnalyzer received None, treating as empty string")
            text = ""
        elif not isinstance(text, str):
            logger.warning(f"MarkdownAnalyzer received {type(text).__name__}, converting to string")
            text = str(text)
        self.text = text
        self.lines = text.split('\n')
        self._headings_cache: dict[str, int] | None = None
        self._word_count_cache: int | None = None
        logger.debug(f"MarkdownAnalyzer initialized with {len(self.lines)} lines")

    def analyze(self) -> dict[str, Any]:
        """Perform comprehensive analysis of the markdown document.

        Returns:
            dict: Analysis results containing:
                - word_count (int): Number of words (excluding code blocks)
                - char_count (int): Total character count
                - line_count (int): Number of lines
                - reading_time (int): Estimated reading time in minutes
                - headings (dict): Count of each heading level (h1-h6)
                - links (int): Number of links
                - images (int): Number of images
                - code_blocks (int): Number of fenced code blocks
                - lists (int): Number of list items
                - blockquotes (int): Number of blockquote lines
                - tables (int): Number of tables
                - readability_score (int): Readability score (0-100)
                - structure_quality (str): Quality rating string
                - broken_links (list): List of detected issues
        """
        logger.debug("Starting document analysis")

        result = {
            'word_count': self._count_words(),
            'char_count': len(self.text),
            'line_count': len(self.lines),
            'reading_time': self._estimate_reading_time(),
            'headings': self._analyze_headings(),
            'links': self._analyze_links(),
            'images': self._count_images(),
            'code_blocks': self._count_code_blocks(),
            'lists': self._count_lists(),
            'blockquotes': self._count_blockquotes(),
            'tables': self._count_tables(),
            'readability_score': self._calculate_readability(),
            'structure_quality': self._analyze_structure_quality(),
            'broken_links': self._detect_potential_issues(),
        }

        logger.debug(f"Analysis complete: {result['word_count']} words, "
                    f"readability: {result['readability_score']}")

        return result

    def _count_words(self) -> int:
        """Count words in the document, excluding code blocks.

        Returns:
            int: Number of words in the document.
        """
        if self._word_count_cache is not None:
            return self._word_count_cache
        # Remove fenced code blocks
        text_without_code = RE_CODE_BLOCK.sub('', self.text)
        # Remove inline code
        text_without_code = RE_INLINE_CODE.sub('', text_without_code)
        # Count words
        words = RE_WORD.findall(text_without_code)
        self._word_count_cache = len(words)
        return self._word_count_cache

    def _estimate_reading_time(self) -> int:
        """Estimate reading time in minutes.

        Uses the configured words per minute rate (default 200 WPM).

        Returns:
            int: Estimated reading time in minutes (minimum 1).
        """
        words = self._count_words()
        return max(1, round(words / WORDS_PER_MINUTE))

    def _analyze_headings(self) -> dict[str, int]:
        """Analyze heading structure (H1-H6 counts).

        Returns:
            dict: Dictionary with keys 'h1' through 'h6' and their counts.
        """
        if self._headings_cache is not None:
            return self._headings_cache
        headings: dict[str, int] = {
            'h1': 0, 'h2': 0, 'h3': 0, 'h4': 0, 'h5': 0, 'h6': 0
        }
        for line in self.lines:
            match = RE_HEADING.match(line.strip())
            if match:
                level = len(match.group(1))
                headings[f'h{level}'] += 1
        self._headings_cache = headings
        return headings

    def _analyze_links(self) -> int:
        """Count markdown links.

        Returns:
            int: Number of links in the document.
        """
        links = RE_LINK.findall(self.text)
        return len(links)

    def _count_images(self) -> int:
        """Count markdown images.

        Returns:
            int: Number of images in the document.
        """
        images = RE_IMAGE.findall(self.text)
        return len(images)

    def _count_code_blocks(self) -> int:
        """Count fenced code blocks.

        Returns:
            int: Number of fenced code blocks (pairs of ```).
        """
        code_blocks = RE_FENCED_CODE.findall(self.text)
        return len(code_blocks) // 2

    def _count_lists(self) -> int:
        """Count list items (ordered and unordered).

        Returns:
            int: Number of list items.
        """
        list_items = 0
        for line in self.lines:
            if RE_UNORDERED_LIST.match(line) or RE_ORDERED_LIST.match(line):
                list_items += 1
        return list_items

    def _count_blockquotes(self) -> int:
        """Count blockquote lines.

        Returns:
            int: Number of blockquote lines.
        """
        quotes = sum(1 for line in self.lines if line.strip().startswith('>'))
        return quotes

    def _count_tables(self) -> int:
        """Count markdown tables (heuristic based).

        Detects tables by looking for lines starting with pipe characters.

        Returns:
            int: Number of tables detected.
        """
        in_table = False
        table_count = 0

        for line in self.lines:
            if '|' in line and line.strip().startswith('|'):
                if not in_table:
                    table_count += 1
                    in_table = True
            else:
                if in_table and not line.strip().startswith('|'):
                    in_table = False

        return table_count

    def _calculate_readability(self) -> int:
        """Calculate the Flesch Reading Ease score (0-100).

        Uses the formally specified Flesch Reading Ease formula (Flesch 1948)
        via the textstat library. Markdown syntax is stripped before scoring
        so that structural markers do not distort the prose-level metric.
        Returns 0 for empty documents.

        Returns:
            int: Flesch Reading Ease score clamped to [0, 100].
                 Higher scores indicate more readable text.
        """
        if self._count_words() == 0:
            return 0

        # Strip Markdown syntax to produce plain prose for scoring.
        clean = RE_CODE_BLOCK.sub(' ', self.text)
        clean = RE_INLINE_CODE.sub(' ', clean)
        # Images before links (images start with !)
        clean = RE_IMAGE.sub(' ', clean)
        # Links — keep display text, drop URL
        clean = RE_LINK.sub(r'\1', clean)
        # Heading markers
        clean = re.sub(r'^#{1,6}\s+', '', clean, flags=re.MULTILINE)
        # Bold / italic markers
        clean = RE_BOLD.sub('', clean)
        clean = RE_ITALIC.sub('', clean)

        score = textstat.flesch_reading_ease(clean.strip())
        return max(0, min(100, round(score)))

    def _analyze_structure_quality(self) -> str:
        """Determine a coarse structure quality rating.

        Returns:
            str: Quality rating - one of:
                - "Excellent": Single H1 with H2 subsections
                - "Good": Some headings present
                - "No structure": No headings
                - "Multiple H1s detected": More than one H1 heading
                - "Needs improvement": Other cases
        """
        headings = self._analyze_headings()
        total_headings = sum(headings.values())

        if total_headings == 0:
            return "No structure"
        elif headings['h1'] > 1:
            return "Multiple H1s detected"
        elif headings['h1'] == 1 and headings['h2'] > 0:
            return "Excellent"
        elif total_headings > 0:
            return "Good"
        else:
            return "Needs improvement"

    def _detect_potential_issues(self) -> list[str]:
        """Detect potential issues in the document.

        Checks for:
        - Empty links (links without URLs)
        - Duplicate headings
        - Very long lines (non-table lines over 120 chars)

        Returns:
            list: List of issue description strings.
        """
        issues: list[str] = []

        # Check for empty links
        empty_links = RE_EMPTY_LINK.findall(self.text)
        if empty_links:
            issues.append(f"{len(empty_links)} empty link(s)")

        # Check for duplicate headings
        headings_text: list[str] = []
        for line in self.lines:
            match = RE_HEADING.match(line.strip())
            if match:
                headings_text.append(match.group(2))

        duplicates = [h for h, count in Counter(headings_text).items() if count > 1]
        if duplicates:
            issues.append(f"{len(duplicates)} duplicate heading(s)")

        # Check for very long lines (excluding tables)
        long_lines = sum(
            1 for line in self.lines
            if len(line) > MAX_LINE_LENGTH_WARNING and not line.strip().startswith('|')
        )
        if long_lines > 5:
            issues.append(f"{long_lines} very long lines")

        return issues if issues else ["No issues detected"]


def get_readability_color(score: int) -> str:
    """Get the color for a readability score.

    Args:
        score: Readability score (0-100).

    Returns:
        str: CSS color string ('green', 'orange', or 'red').
    """
    if score >= READABILITY_GREEN:
        return "green"
    elif score >= READABILITY_ORANGE:
        return "orange"
    else:
        return "red"


def get_structure_color(quality: str) -> str:
    """Get the color for a structure quality rating.

    Args:
        quality: Structure quality string from _analyze_structure_quality().

    Returns:
        str: CSS color string ('green', 'orange', or 'red').
    """
    if quality == "Excellent":
        return "green"
    elif quality == "Good":
        return "orange"
    else:
        return "red"
