"""Smart Assistant panel widget.

Displays live document statistics and quality metrics derived from
:class:`src.core.analyzer.MarkdownAnalyzer`.
"""

from __future__ import annotations

from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QGroupBox,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class AssistantPanel(QWidget):
    """Side-panel that shows document statistics, structure and quality info."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._build_ui()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def update_metrics(self, metrics: dict) -> None:
        """Refresh every label from a metrics dict produced by MarkdownAnalyzer."""
        # Statistics
        self._stats_labels["words"].setText(f"Words: {metrics['word_count']}")
        self._stats_labels["chars"].setText(f"Characters: {metrics['char_count']}")
        self._stats_labels["lines"].setText(f"Lines: {metrics['line_count']}")
        self._stats_labels["reading_time"].setText(
            f"Reading time: {metrics['reading_time']} min"
        )

        # Structure
        total_headings = sum(metrics["headings"].values())
        breakdown = ", ".join(
            f"H{i}: {metrics['headings'][f'h{i}']}"
            for i in range(1, 7)
            if metrics["headings"][f"h{i}"] > 0
        )
        heading_text = (
            f"Headings: {total_headings} ({breakdown})"
            if breakdown
            else f"Headings: {total_headings}"
        )
        self._structure_labels["headings"].setText(heading_text)
        self._structure_labels["links"].setText(f"Links: {metrics['links']}")
        self._structure_labels["images"].setText(f"Images: {metrics['images']}")
        self._structure_labels["code_blocks"].setText(
            f"Code blocks: {metrics['code_blocks']}"
        )
        self._structure_labels["lists"].setText(f"List items: {metrics['lists']}")
        self._structure_labels["blockquotes"].setText(
            f"Blockquotes: {metrics['blockquotes']}"
        )
        self._structure_labels["tables"].setText(f"Tables: {metrics['tables']}")

        # Quality
        readability = metrics["readability_score"]
        r_color = (
            "green" if readability >= 80 else "orange" if readability >= 60 else "red"
        )
        self._quality_labels["readability"].setText(f"Readability: {readability}/100")
        self._quality_labels["readability"].setStyleSheet(
            f"color: {r_color}; font-weight: bold;"
        )

        structure_quality = metrics["structure_quality"]
        s_color = (
            "green"
            if structure_quality == "Excellent"
            else "orange"
            if structure_quality == "Good"
            else "red"
        )
        self._quality_labels["structure_quality"].setText(
            f"Structure: {structure_quality}"
        )
        self._quality_labels["structure_quality"].setStyleSheet(
            f"color: {s_color}; font-weight: bold;"
        )

        # Issues
        self._issues_label.setText(
            "\n".join(f"• {issue}" for issue in metrics["broken_links"])
        )

    def clear(self) -> None:
        """Reset all labels to their default empty-document values."""
        self._stats_labels["words"].setText("Words: 0")
        self._stats_labels["chars"].setText("Characters: 0")
        self._stats_labels["lines"].setText("Lines: 0")
        self._stats_labels["reading_time"].setText("Reading time: 1 min")

        self._structure_labels["headings"].setText("Headings: 0")
        self._structure_labels["links"].setText("Links: 0")
        self._structure_labels["images"].setText("Images: 0")
        self._structure_labels["code_blocks"].setText("Code blocks: 0")
        self._structure_labels["lists"].setText("List items: 0")
        self._structure_labels["blockquotes"].setText("Blockquotes: 0")
        self._structure_labels["tables"].setText("Tables: 0")

        self._quality_labels["readability"].setText("Readability: --")
        self._quality_labels["readability"].setStyleSheet("")
        self._quality_labels["structure_quality"].setText("Structure: --")
        self._quality_labels["structure_quality"].setStyleSheet("")

        self._issues_label.setText("No issues detected")

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        layout = QVBoxLayout()

        # Header
        title = QLabel("Smart Assistant")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Statistics group
        stats_group = QGroupBox("Document Statistics")
        stats_inner = QVBoxLayout()
        self._stats_labels: dict[str, QLabel] = {
            "words": QLabel("Words: 0"),
            "chars": QLabel("Characters: 0"),
            "lines": QLabel("Lines: 0"),
            "reading_time": QLabel("Reading time: 1 min"),
        }
        for lbl in self._stats_labels.values():
            stats_inner.addWidget(lbl)
        stats_group.setLayout(stats_inner)
        layout.addWidget(stats_group)

        # Structure group
        structure_group = QGroupBox("Document Structure")
        structure_inner = QVBoxLayout()
        self._structure_labels: dict[str, QLabel] = {
            "headings": QLabel("Headings: 0"),
            "links": QLabel("Links: 0"),
            "images": QLabel("Images: 0"),
            "code_blocks": QLabel("Code blocks: 0"),
            "lists": QLabel("List items: 0"),
            "blockquotes": QLabel("Blockquotes: 0"),
            "tables": QLabel("Tables: 0"),
        }
        for lbl in self._structure_labels.values():
            structure_inner.addWidget(lbl)
        structure_group.setLayout(structure_inner)
        layout.addWidget(structure_group)

        # Quality group
        quality_group = QGroupBox("Quality Analysis")
        quality_inner = QVBoxLayout()
        self._quality_labels: dict[str, QLabel] = {
            "readability": QLabel("Readability: --"),
            "structure_quality": QLabel("Structure: --"),
        }
        for lbl in self._quality_labels.values():
            quality_inner.addWidget(lbl)
        quality_group.setLayout(quality_inner)
        layout.addWidget(quality_group)

        # Issues group
        issues_group = QGroupBox("Potential Issues")
        issues_inner = QVBoxLayout()
        self._issues_label = QLabel("No issues detected")
        self._issues_label.setWordWrap(True)
        issues_inner.addWidget(self._issues_label)
        issues_group.setLayout(issues_inner)
        layout.addWidget(issues_group)

        # Format button — the parent (MainWindow) connects the signal externally
        self.format_button = QPushButton("Auto-Format Document")
        layout.addWidget(self.format_button)

        layout.addStretch()
        self.setLayout(layout)
