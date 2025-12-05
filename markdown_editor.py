#!/usr/bin/env python3
"""
Smart Markdown Editor with Live Preview and Smart Assistant
A modern, cross-platform desktop markdown editor built with PySide6.
"""

import sys
import os
import re
import markdown
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QTextEdit, QSplitter, QMenuBar,
                               QMenu, QFileDialog, QMessageBox, QLabel,
                               QGroupBox, QScrollArea, QPushButton, QDockWidget)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QColor, QPalette
from PySide6.QtWebEngineWidgets import QWebEngineView
from collections import Counter
from datetime import datetime

# Additional imports for export functionality
try:
    from docx import Document
    from docx.shared import Inches
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import weasyprint
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

try:
    import html2text
    HTML2TEXT_AVAILABLE = True
except ImportError:
    HTML2TEXT_AVAILABLE = False


class MarkdownAnalyzer:
    """Smart Markdown Document Analyzer - analyzes quality, structure, and provides insights."""

    def __init__(self, text):
        self.text = text
        self.lines = text.split('\n')

    def analyze(self):
        """Perform comprehensive analysis and return metrics."""
        return {
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

    def _count_words(self):
        """Count words in the document."""
        # Remove code blocks first
        text_without_code = re.sub(r'```.*?```', '', self.text, flags=re.DOTALL)
        # Remove inline code
        text_without_code = re.sub(r'`[^`]+`', '', text_without_code)
        # Count words
        words = re.findall(r'\b\w+\b', text_without_code)
        return len(words)

    def _estimate_reading_time(self):
        """Estimate reading time in minutes (average 200 words/min)."""
        words = self._count_words()
        return max(1, round(words / 200))

    def _analyze_headings(self):
        """Analyze heading structure."""
        headings = {'h1': 0, 'h2': 0, 'h3': 0, 'h4': 0, 'h5': 0, 'h6': 0}
        for line in self.lines:
            match = re.match(r'^(#{1,6})\s+(.+)$', line.strip())
            if match:
                level = len(match.group(1))
                headings[f'h{level}'] += 1
        return headings

    def _analyze_links(self):
        """Count and analyze links."""
        # Markdown links: [text](url)
        links = re.findall(r'\[([^\]]+)\]\(([^\)]+)\)', self.text)
        return len(links)

    def _count_images(self):
        """Count images in the document."""
        # Markdown images: ![alt](url)
        images = re.findall(r'!\[([^\]]*)\]\(([^\)]+)\)', self.text)
        return len(images)

    def _count_code_blocks(self):
        """Count code blocks."""
        code_blocks = re.findall(r'```', self.text)
        return len(code_blocks) // 2  # Each block has opening and closing

    def _count_lists(self):
        """Count list items."""
        list_items = 0
        for line in self.lines:
            if re.match(r'^\s*[-*+]\s+', line) or re.match(r'^\s*\d+\.\s+', line):
                list_items += 1
        return list_items

    def _count_blockquotes(self):
        """Count blockquote lines."""
        quotes = sum(1 for line in self.lines if line.strip().startswith('>'))
        return quotes

    def _count_tables(self):
        """Count tables in the document."""
        # Look for table rows (containing |)
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

    def _calculate_readability(self):
        """Calculate a simple readability score (0-100)."""
        score = 100

        # Penalize very long paragraphs
        paragraphs = self.text.split('\n\n')
        avg_paragraph_length = sum(len(p.split()) for p in paragraphs) / max(len(paragraphs), 1)
        if avg_paragraph_length > 100:
            score -= 10
        elif avg_paragraph_length > 150:
            score -= 20

        # Reward good heading structure
        headings = self._analyze_headings()
        if headings['h1'] >= 1 and headings['h2'] > 0:
            score += 10

        # Penalize documents without any structure
        if sum(headings.values()) == 0 and self._count_words() > 50:
            score -= 15

        return max(0, min(100, score))

    def _analyze_structure_quality(self):
        """Analyze document structure quality."""
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

    def _detect_potential_issues(self):
        """Detect potential issues in the document."""
        issues = []

        # Check for empty links
        empty_links = re.findall(r'\[([^\]]+)\]\(\s*\)', self.text)
        if empty_links:
            issues.append(f"{len(empty_links)} empty link(s)")

        # Check for duplicate headings
        headings_text = []
        for line in self.lines:
            match = re.match(r'^#{1,6}\s+(.+)$', line.strip())
            if match:
                headings_text.append(match.group(1))

        duplicates = [h for h, count in Counter(headings_text).items() if count > 1]
        if duplicates:
            issues.append(f"{len(duplicates)} duplicate heading(s)")

        # Check for very long lines (potential formatting issues)
        long_lines = sum(1 for line in self.lines if len(line) > 120 and not line.strip().startswith('|'))
        if long_lines > 5:
            issues.append(f"{long_lines} very long lines")

        return issues if issues else ["No issues detected"]


class MarkdownEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Markdown Editor - Smart Assistant")
        self.setGeometry(100, 100, 1400, 900)

        # Initialize UI
        self.init_ui()

        # Set up timer for live preview updates
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self.update_preview)

        # Set up timer for analysis updates
        self.analysis_timer = QTimer()
        self.analysis_timer.setSingleShot(True)
        self.analysis_timer.timeout.connect(self.update_analysis)

        # Initial preview and analysis update
        self.update_preview()
        self.update_analysis()
    
    def init_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Create main horizontal splitter (editor | preview | assistant)
        main_splitter = QSplitter(Qt.Horizontal)

        # Create text editor for markdown input
        self.editor = QTextEdit()
        self.editor.setPlaceholderText("Type your markdown here...")
        self.editor.textChanged.connect(self.on_text_changed)
        
        # Set editor styling for consistent appearance
        self.editor.setStyleSheet("""
            QTextEdit {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #ddd;
                font-family: 'SF Mono', 'Monaco', 'Menlo', 'Consolas', monospace;
                font-size: 14px;
                padding: 10px;
                selection-background-color: #0078d4;
                selection-color: #ffffff;
            }
        """)

        # Create web view for HTML preview
        self.preview = QWebEngineView()

        # Create Smart Assistant Panel
        self.assistant_panel = self.create_assistant_panel()

        # Add widgets to splitter
        main_splitter.addWidget(self.editor)
        main_splitter.addWidget(self.preview)
        main_splitter.addWidget(self.assistant_panel)

        # Set splitter sizes (30% editor, 45% preview, 25% assistant)
        main_splitter.setSizes([420, 630, 350])

        # Create main layout
        layout = QVBoxLayout()
        layout.addWidget(main_splitter)
        central_widget.setLayout(layout)

        # Create menu bar
        self.create_menu_bar()

    def create_assistant_panel(self):
        """Create the Smart Markdown Assistant panel."""
        panel = QWidget()
        panel_layout = QVBoxLayout()

        # Title
        title = QLabel("Smart Assistant")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        panel_layout.addWidget(title)

        # Statistics Group
        stats_group = QGroupBox("Document Statistics")
        stats_layout = QVBoxLayout()

        self.stats_labels = {
            'words': QLabel("Words: 0"),
            'chars': QLabel("Characters: 0"),
            'lines': QLabel("Lines: 0"),
            'reading_time': QLabel("Reading time: 1 min"),
        }

        for label in self.stats_labels.values():
            stats_layout.addWidget(label)

        stats_group.setLayout(stats_layout)
        panel_layout.addWidget(stats_group)

        # Structure Analysis Group
        structure_group = QGroupBox("Document Structure")
        structure_layout = QVBoxLayout()

        self.structure_labels = {
            'headings': QLabel("Headings: 0"),
            'links': QLabel("Links: 0"),
            'images': QLabel("Images: 0"),
            'code_blocks': QLabel("Code blocks: 0"),
            'lists': QLabel("List items: 0"),
            'blockquotes': QLabel("Blockquotes: 0"),
            'tables': QLabel("Tables: 0"),
        }

        for label in self.structure_labels.values():
            structure_layout.addWidget(label)

        structure_group.setLayout(structure_layout)
        panel_layout.addWidget(structure_group)

        # Quality Analysis Group
        quality_group = QGroupBox("Quality Analysis")
        quality_layout = QVBoxLayout()

        self.quality_labels = {
            'readability': QLabel("Readability: --"),
            'structure_quality': QLabel("Structure: --"),
        }

        for label in self.quality_labels.values():
            quality_layout.addWidget(label)

        quality_group.setLayout(quality_layout)
        panel_layout.addWidget(quality_group)

        # Issues Group
        issues_group = QGroupBox("Potential Issues")
        issues_layout = QVBoxLayout()

        self.issues_label = QLabel("No issues detected")
        self.issues_label.setWordWrap(True)
        issues_layout.addWidget(self.issues_label)

        issues_group.setLayout(issues_layout)
        panel_layout.addWidget(issues_group)

        # Auto-format button
        self.format_button = QPushButton("Auto-Format Document")
        self.format_button.clicked.connect(self.auto_format_document)
        panel_layout.addWidget(self.format_button)

        # Add stretch to push everything to top
        panel_layout.addStretch()

        panel.setLayout(panel_layout)
        return panel
    
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        # New action
        new_action = file_menu.addAction("New")
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_file)
        
        # Open action
        open_action = file_menu.addAction("Open")
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        
        # Save action
        save_action = file_menu.addAction("Save")
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        
        # Save As action
        save_as_action = file_menu.addAction("Save As...")
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_file_as)
        
        file_menu.addSeparator()
        
        # Export submenu
        export_menu = file_menu.addMenu("Export As")
        
        # Export actions
        export_md_action = export_menu.addAction("Markdown (.md)")
        export_md_action.triggered.connect(lambda: self.export_file('md'))
        
        export_txt_action = export_menu.addAction("Plain Text (.txt)")
        export_txt_action.triggered.connect(lambda: self.export_file('txt'))
        
        export_html_action = export_menu.addAction("HTML (.html)")
        export_html_action.triggered.connect(lambda: self.export_file('html'))
        
        if DOCX_AVAILABLE:
            export_docx_action = export_menu.addAction("Word Document (.docx)")
            export_docx_action.triggered.connect(lambda: self.export_file('docx'))
        
        if PDF_AVAILABLE:
            export_pdf_action = export_menu.addAction("PDF Document (.pdf)")
            export_pdf_action.triggered.connect(lambda: self.export_file('pdf'))
        
        export_rtf_action = export_menu.addAction("Rich Text Format (.rtf)")
        export_rtf_action.triggered.connect(lambda: self.export_file('rtf'))
        
        export_odt_action = export_menu.addAction("OpenDocument Text (.odt)")
        export_odt_action.triggered.connect(lambda: self.export_file('odt'))
        
        file_menu.addSeparator()
        
        # Exit action
        exit_action = file_menu.addAction("Exit")
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        
        # Edit menu
        edit_menu = menubar.addMenu("Edit")
        
        # Undo action
        undo_action = edit_menu.addAction("Undo")
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.editor.undo)
        
        # Redo action
        redo_action = edit_menu.addAction("Redo")
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.editor.redo)
    
    def on_text_changed(self):
        # Start timer to update preview after a short delay
        # This prevents excessive updates while typing
        self.update_timer.start(300)  # 300ms delay
        # Also update analysis with a longer delay
        self.analysis_timer.start(800)  # 800ms delay for analysis
    
    def update_preview(self):
        # Get markdown text from editor
        markdown_text = self.editor.toPlainText()
        
        # Convert markdown to HTML
        html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
        
        # Add basic CSS styling
        styled_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    line-height: 1.6;
                    color: #333;
                    max-width: 800px;
                    margin: 0 auto;
                    padding: 20px;
                    background-color: #fff;
                }}
                h1, h2, h3, h4, h5, h6 {{
                    margin-top: 24px;
                    margin-bottom: 16px;
                    font-weight: 600;
                    line-height: 1.25;
                }}
                h1 {{ font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
                h2 {{ font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
                h3 {{ font-size: 1.25em; }}
                h4 {{ font-size: 1em; }}
                h5 {{ font-size: 0.875em; }}
                h6 {{ font-size: 0.85em; color: #6a737d; }}
                p {{ margin-bottom: 16px; }}
                code {{
                    background-color: #f6f8fa;
                    border-radius: 3px;
                    font-size: 85%;
                    margin: 0;
                    padding: 0.2em 0.4em;
                }}
                pre {{
                    background-color: #f6f8fa;
                    border-radius: 6px;
                    padding: 16px;
                    overflow: auto;
                    font-size: 85%;
                    line-height: 1.45;
                }}
                pre code {{
                    background-color: transparent;
                    border: 0;
                    display: inline;
                    line-height: inherit;
                    margin: 0;
                    max-width: auto;
                    overflow: visible;
                    padding: 0;
                    word-wrap: normal;
                }}
                blockquote {{
                    border-left: 0.25em solid #dfe2e5;
                    color: #6a737d;
                    padding: 0 1em;
                    margin: 0 0 16px 0;
                }}
                table {{
                    border-spacing: 0;
                    border-collapse: collapse;
                    margin-bottom: 16px;
                }}
                table th, table td {{
                    border: 1px solid #dfe2e5;
                    padding: 6px 13px;
                }}
                table th {{
                    background-color: #f6f8fa;
                    font-weight: 600;
                }}
                table tr:nth-child(2n) {{
                    background-color: #f6f8fa;
                }}
                ul, ol {{
                    padding-left: 2em;
                    margin-bottom: 16px;
                }}
                li {{
                    margin-bottom: 0.25em;
                }}
                a {{
                    color: #0366d6;
                    text-decoration: none;
                }}
                a:hover {{
                    text-decoration: underline;
                }}
                img {{
                    max-width: 100%;
                    height: auto;
                }}
                hr {{
                    border: none;
                    border-top: 1px solid #eaecef;
                    height: 1px;
                    margin: 24px 0;
                }}
            </style>
        </head>
        <body>
            {html}
        </body>
        </html>
        """
        
        # Set HTML to preview
        self.preview.setHtml(styled_html)

    def update_analysis(self):
        """Update the Smart Assistant panel with document analysis."""
        markdown_text = self.editor.toPlainText()

        if not markdown_text.strip():
            # Reset all labels for empty document
            self.stats_labels['words'].setText("Words: 0")
            self.stats_labels['chars'].setText("Characters: 0")
            self.stats_labels['lines'].setText("Lines: 0")
            self.stats_labels['reading_time'].setText("Reading time: 1 min")

            self.structure_labels['headings'].setText("Headings: 0")
            self.structure_labels['links'].setText("Links: 0")
            self.structure_labels['images'].setText("Images: 0")
            self.structure_labels['code_blocks'].setText("Code blocks: 0")
            self.structure_labels['lists'].setText("List items: 0")
            self.structure_labels['blockquotes'].setText("Blockquotes: 0")
            self.structure_labels['tables'].setText("Tables: 0")

            self.quality_labels['readability'].setText("Readability: --")
            self.quality_labels['structure_quality'].setText("Structure: --")
            self.issues_label.setText("No issues detected")
            return

        # Analyze document
        analyzer = MarkdownAnalyzer(markdown_text)
        metrics = analyzer.analyze()

        # Update statistics
        self.stats_labels['words'].setText(f"Words: {metrics['word_count']}")
        self.stats_labels['chars'].setText(f"Characters: {metrics['char_count']}")
        self.stats_labels['lines'].setText(f"Lines: {metrics['line_count']}")
        self.stats_labels['reading_time'].setText(f"Reading time: {metrics['reading_time']} min")

        # Update structure
        total_headings = sum(metrics['headings'].values())
        heading_breakdown = ', '.join([f"H{i}: {metrics['headings'][f'h{i}']}"
                                      for i in range(1, 7) if metrics['headings'][f'h{i}'] > 0])
        self.structure_labels['headings'].setText(f"Headings: {total_headings} ({heading_breakdown})" if heading_breakdown else f"Headings: {total_headings}")
        self.structure_labels['links'].setText(f"Links: {metrics['links']}")
        self.structure_labels['images'].setText(f"Images: {metrics['images']}")
        self.structure_labels['code_blocks'].setText(f"Code blocks: {metrics['code_blocks']}")
        self.structure_labels['lists'].setText(f"List items: {metrics['lists']}")
        self.structure_labels['blockquotes'].setText(f"Blockquotes: {metrics['blockquotes']}")
        self.structure_labels['tables'].setText(f"Tables: {metrics['tables']}")

        # Update quality analysis
        readability = metrics['readability_score']
        readability_color = "green" if readability >= 80 else "orange" if readability >= 60 else "red"
        self.quality_labels['readability'].setText(f"Readability: {readability}/100")
        self.quality_labels['readability'].setStyleSheet(f"color: {readability_color}; font-weight: bold;")

        structure_quality = metrics['structure_quality']
        structure_color = "green" if structure_quality == "Excellent" else "orange" if structure_quality == "Good" else "red"
        self.quality_labels['structure_quality'].setText(f"Structure: {structure_quality}")
        self.quality_labels['structure_quality'].setStyleSheet(f"color: {structure_color}; font-weight: bold;")

        # Update issues
        issues_text = "\n".join(f"• {issue}" for issue in metrics['broken_links'])
        self.issues_label.setText(issues_text)

    def auto_format_document(self):
        """Auto-format the markdown document with best practices."""
        markdown_text = self.editor.toPlainText()

        if not markdown_text.strip():
            QMessageBox.information(self, "Auto-Format", "Document is empty. Nothing to format.")
            return

        # Format the document
        formatted_text = self.format_markdown(markdown_text)

        # Update editor
        self.editor.setPlainText(formatted_text)

        QMessageBox.information(self, "Auto-Format", "Document formatted successfully!")

    def format_markdown(self, text):
        """Apply auto-formatting rules to markdown text."""
        lines = text.split('\n')
        formatted_lines = []
        prev_was_heading = False
        prev_was_empty = False

        for i, line in enumerate(lines):
            stripped = line.strip()

            # Add blank line before headings (except first line)
            if stripped.startswith('#') and i > 0 and not prev_was_empty:
                formatted_lines.append('')

            # Ensure space after # in headings
            if stripped.startswith('#'):
                match = re.match(r'^(#{1,6})(\S)', stripped)
                if match:
                    level = match.group(1)
                    rest = stripped[len(level):]
                    formatted_lines.append(f"{level} {rest}")
                    prev_was_heading = True
                    prev_was_empty = False
                    continue

            # Ensure space after list markers
            if re.match(r'^(\s*)([-*+])(\S)', line):
                match = re.match(r'^(\s*)([-*+])(.*)$', line)
                indent = match.group(1)
                marker = match.group(2)
                content = match.group(3).strip()
                formatted_lines.append(f"{indent}{marker} {content}")
                prev_was_heading = False
                prev_was_empty = False
                continue

            # Ensure space after numbered list markers
            if re.match(r'^(\s*)(\d+\.)(\S)', line):
                match = re.match(r'^(\s*)(\d+\.)(.*)$', line)
                indent = match.group(1)
                marker = match.group(2)
                content = match.group(3).strip()
                formatted_lines.append(f"{indent}{marker} {content}")
                prev_was_heading = False
                prev_was_empty = False
                continue

            # Track empty lines
            if not stripped:
                # Don't add multiple consecutive empty lines
                if not prev_was_empty:
                    formatted_lines.append('')
                    prev_was_empty = True
            else:
                formatted_lines.append(line)
                prev_was_empty = False
                prev_was_heading = False

        # Remove trailing empty lines
        while formatted_lines and not formatted_lines[-1]:
            formatted_lines.pop()

        # Ensure file ends with single newline
        return '\n'.join(formatted_lines)

    def new_file(self):
        self.editor.clear()
        self.current_file = None
    
    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Markdown File", "", "Markdown Files (*.md);;All Files (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.editor.setPlainText(content)
                    self.current_file = file_path
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not open file: {str(e)}")
    
    def save_file(self):
        if hasattr(self, 'current_file') and self.current_file:
            self.save_to_file(self.current_file)
        else:
            self.save_file_as()
    
    def save_file_as(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Markdown File", "", "Markdown Files (*.md);;All Files (*)"
        )
        
        if file_path:
            self.save_to_file(file_path)
    
    def save_to_file(self, file_path):
        try:
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(self.editor.toPlainText())
                self.current_file = file_path
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save file: {str(e)}")
    
    def export_file(self, file_format):
        """Export the current markdown content to various formats."""
        file_filters = {
            'md': 'Markdown Files (*.md)',
            'txt': 'Text Files (*.txt)',
            'html': 'HTML Files (*.html)',
            'docx': 'Word Documents (*.docx)',
            'pdf': 'PDF Files (*.pdf)',
            'rtf': 'Rich Text Format (*.rtf)',
            'odt': 'OpenDocument Text (*.odt)'
        }
        
        if file_format not in file_filters:
            QMessageBox.warning(self, "Export Error", f"Unsupported export format: {file_format}")
            return
        
        # Check if required libraries are available
        if file_format == 'docx' and not DOCX_AVAILABLE:
            QMessageBox.warning(self, "Export Error", "python-docx library is not installed. Please install it with: pip install python-docx")
            return
        
        if file_format == 'pdf' and not (PDF_AVAILABLE or WEASYPRINT_AVAILABLE):
            QMessageBox.warning(self, "Export Error", "Neither reportlab nor weasyprint is installed. Please install one with: pip install reportlab or pip install weasyprint")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, f"Export as {file_format.upper()}", "", file_filters[file_format]
        )
        
        if file_path:
            try:
                if file_format == 'md':
                    self.export_as_markdown(file_path)
                elif file_format == 'txt':
                    self.export_as_text(file_path)
                elif file_format == 'html':
                    self.export_as_html(file_path)
                elif file_format == 'docx':
                    self.export_as_docx(file_path)
                elif file_format == 'pdf':
                    self.export_as_pdf(file_path)
                elif file_format == 'rtf':
                    self.export_as_rtf(file_path)
                elif file_format == 'odt':
                    self.export_as_odt(file_path)
                
                QMessageBox.information(self, "Export Successful", f"File exported successfully to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Could not export file: {str(e)}")
    
    def export_as_markdown(self, file_path):
        """Export as markdown file."""
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(self.editor.toPlainText())
    
    def export_as_text(self, file_path):
        """Export as plain text file."""
        markdown_text = self.editor.toPlainText()
        # Simple conversion: remove markdown syntax for plain text
        import re
        text = markdown_text
        # Remove headers
        text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
        # Remove bold and italic
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        text = re.sub(r'\*(.*?)\*', r'\1', text)
        # Remove code blocks
        text = re.sub(r'```.*?\n(.*?)\n```', r'\1', text, flags=re.DOTALL)
        text = re.sub(r'`(.*?)`', r'\1', text)
        # Remove links but keep text
        text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
        # Remove images
        text = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', r'[\1]', text)
        
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(text)
    
    def export_as_html(self, file_path):
        """Export as HTML file."""
        markdown_text = self.editor.toPlainText()
        html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
        
        styled_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Exported Markdown Document</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #fff;
        }}
        h1, h2, h3, h4, h5, h6 {{
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}
        h1 {{ font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
        h2 {{ font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
        h3 {{ font-size: 1.25em; }}
        h4 {{ font-size: 1em; }}
        h5 {{ font-size: 0.875em; }}
        h6 {{ font-size: 0.85em; color: #6a737d; }}
        p {{ margin-bottom: 16px; }}
        code {{
            background-color: #f6f8fa;
            border-radius: 3px;
            font-size: 85%;
            margin: 0;
            padding: 0.2em 0.4em;
        }}
        pre {{
            background-color: #f6f8fa;
            border-radius: 6px;
            padding: 16px;
            overflow: auto;
            font-size: 85%;
            line-height: 1.45;
        }}
        pre code {{
            background-color: transparent;
            border: 0;
            display: inline;
            line-height: inherit;
            margin: 0;
            max-width: auto;
            overflow: visible;
            padding: 0;
            word-wrap: normal;
        }}
        blockquote {{
            border-left: 0.25em solid #dfe2e5;
            color: #6a737d;
            padding: 0 1em;
            margin: 0 0 16px 0;
        }}
        table {{
            border-spacing: 0;
            border-collapse: collapse;
            margin-bottom: 16px;
        }}
        table th, table td {{
            border: 1px solid #dfe2e5;
            padding: 6px 13px;
        }}
        table th {{
            background-color: #f6f8fa;
            font-weight: 600;
        }}
        table tr:nth-child(2n) {{
            background-color: #f6f8fa;
        }}
        ul, ol {{
            padding-left: 2em;
            margin-bottom: 16px;
        }}
        li {{
            margin-bottom: 0.25em;
        }}
        a {{
            color: #0366d6;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}
        img {{
            max-width: 100%;
            height: auto;
        }}
        hr {{
            border: none;
            border-top: 1px solid #eaecef;
            height: 1px;
            margin: 24px 0;
        }}
    </style>
</head>
<body>
    {html}
</body>
</html>"""
        
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(styled_html)
    
    def export_as_docx(self, file_path):
        """Export as Word document."""
        if not DOCX_AVAILABLE:
            raise ImportError("python-docx library is not available")
        
        markdown_text = self.editor.toPlainText()
        html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
        
        # Create a new Word document
        doc = Document()
        
        # Add title
        doc.add_heading('Exported Markdown Document', 0)
        
        # Parse and add content
        lines = markdown_text.split('\n')
        in_code_block = False
        code_content = []
        
        for line in lines:
            if line.strip().startswith('```'):
                if in_code_block:
                    # End of code block
                    if code_content:
                        paragraph = doc.add_paragraph()
                        run = paragraph.add_run('\n'.join(code_content))
                        run.font.name = 'Courier New'
                        code_content = []
                    in_code_block = False
                else:
                    # Start of code block
                    in_code_block = True
                continue
            
            if in_code_block:
                code_content.append(line)
                continue
            
            # Headers
            if line.startswith('# '):
                doc.add_heading(line[2:], level=1)
            elif line.startswith('## '):
                doc.add_heading(line[3:], level=2)
            elif line.startswith('### '):
                doc.add_heading(line[4:], level=3)
            elif line.startswith('#### '):
                doc.add_heading(line[5:], level=4)
            elif line.startswith('##### '):
                doc.add_heading(line[6:], level=5)
            elif line.startswith('###### '):
                doc.add_heading(line[7:], level=6)
            # Horizontal rule
            elif line.strip() == '---' or line.strip() == '***':
                doc.add_paragraph('_' * 50)
            # Lists
            elif line.strip().startswith(('- ', '* ', '+ ')):
                p = doc.add_paragraph(line.strip()[2:], style='List Bullet')
            elif line.strip() and line.strip()[0].isdigit() and line.strip().find('.') == 1:
                p = doc.add_paragraph(line.strip()[3:], style='List Number')
            # Empty line
            elif not line.strip():
                doc.add_paragraph()
            # Regular paragraph
            elif line.strip():
                # Simple markdown processing
                processed_line = line
                # Bold
                processed_line = re.sub(r'\*\*(.*?)\*\*', r'\1', processed_line)
                # Italic
                processed_line = re.sub(r'\*(.*?)\*', r'\1', processed_line)
                # Code
                processed_line = re.sub(r'`(.*?)`', r'\1', processed_line)
                # Links
                processed_line = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', processed_line)
                
                doc.add_paragraph(processed_line)
        
        doc.save(file_path)
    
    def export_as_pdf(self, file_path):
        """Export as PDF document."""
        markdown_text = self.editor.toPlainText()
        
        # Try WeasyPrint first (better quality)
        if WEASYPRINT_AVAILABLE:
            html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
            
            styled_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Exported Markdown Document</title>
    <style>
        @page {{
            size: letter;
            margin: 1in;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #333;
            font-size: 12pt;
        }}
        h1, h2, h3, h4, h5, h6 {{
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}
        h1 {{ font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
        h2 {{ font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; }}
        h3 {{ font-size: 1.25em; }}
        h4 {{ font-size: 1em; }}
        h5 {{ font-size: 0.875em; }}
        h6 {{ font-size: 0.85em; color: #6a737d; }}
        p {{ margin-bottom: 16px; }}
        code {{
            background-color: #f6f8fa;
            border-radius: 3px;
            font-size: 85%;
            margin: 0;
            padding: 0.2em 0.4em;
        }}
        pre {{
            background-color: #f6f8fa;
            border-radius: 6px;
            padding: 16px;
            overflow: auto;
            font-size: 85%;
            line-height: 1.45;
        }}
        pre code {{
            background-color: transparent;
            border: 0;
            display: inline;
            line-height: inherit;
            margin: 0;
            max-width: auto;
            overflow: visible;
            padding: 0;
            word-wrap: normal;
        }}
        blockquote {{
            border-left: 0.25em solid #dfe2e5;
            color: #6a737d;
            padding: 0 1em;
            margin: 0 0 16px 0;
        }}
        table {{
            border-spacing: 0;
            border-collapse: collapse;
            margin-bottom: 16px;
        }}
        table th, table td {{
            border: 1px solid #dfe2e5;
            padding: 6px 13px;
        }}
        table th {{
            background-color: #f6f8fa;
            font-weight: 600;
        }}
        table tr:nth-child(2n) {{
            background-color: #f6f8fa;
        }}
        ul, ol {{
            padding-left: 2em;
            margin-bottom: 16px;
        }}
        li {{
            margin-bottom: 0.25em;
        }}
        a {{
            color: #0366d6;
            text-decoration: none;
        }}
        img {{
            max-width: 100%;
            height: auto;
        }}
        hr {{
            border: none;
            border-top: 1px solid #eaecef;
            height: 1px;
            margin: 24px 0;
        }}
    </style>
</head>
<body>
    {html}
</body>
</html>"""
            
            # Generate PDF
            weasyprint.HTML(string=styled_html).write_pdf(file_path)
        
        # Fallback to ReportLab
        elif PDF_AVAILABLE:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []
            
            # Add title
            title_style = styles['Title']
            title = Paragraph("Exported Markdown Document", title_style)
            story.append(title)
            story.append(Spacer(1, 12))
            
            # Parse and add content
            lines = markdown_text.split('\n')
            in_code_block = False
            code_content = []
            
            for line in lines:
                if line.strip().startswith('```'):
                    if in_code_block:
                        # End of code block
                        if code_content:
                            code_text = '\n'.join(code_content)
                            code_style = styles['Code']
                            code_para = Paragraph(code_text, code_style)
                            story.append(code_para)
                            story.append(Spacer(1, 6))
                            code_content = []
                        in_code_block = False
                    else:
                        # Start of code block
                        in_code_block = True
                    continue
                
                if in_code_block:
                    code_content.append(line)
                    continue
                
                # Headers
                if line.startswith('# '):
                    story.append(Paragraph(line[2:], styles['Heading1']))
                elif line.startswith('## '):
                    story.append(Paragraph(line[3:], styles['Heading2']))
                elif line.startswith('### '):
                    story.append(Paragraph(line[4:], styles['Heading3']))
                elif line.startswith('#### '):
                    story.append(Paragraph(line[5:], styles['Heading4']))
                elif line.startswith('##### '):
                    story.append(Paragraph(line[6:], styles['Heading5']))
                elif line.startswith('###### '):
                    story.append(Paragraph(line[7:], styles['Heading6']))
                # Horizontal rule
                elif line.strip() == '---' or line.strip() == '***':
                    story.append(Spacer(1, 12))
                # Lists (simplified)
                elif line.strip().startswith(('- ', '* ', '+ ')):
                    story.append(Paragraph(f"• {line.strip()[2:]}", styles['Normal']))
                elif line.strip() and line.strip()[0].isdigit() and line.strip().find('.') == 1:
                    story.append(Paragraph(f"{line.strip()}", styles['Normal']))
                # Empty line
                elif not line.strip():
                    story.append(Spacer(1, 6))
                # Regular paragraph
                elif line.strip():
                    # Simple markdown processing
                    processed_line = line
                    # Bold
                    processed_line = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', processed_line)
                    # Italic
                    processed_line = re.sub(r'\*(.*?)\*', r'<i>\1</i>', processed_line)
                    # Code
                    processed_line = re.sub(r'`(.*?)`', r'<font name="Courier">\1</font>', processed_line)
                    # Links
                    processed_line = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', processed_line)
                    
                    story.append(Paragraph(processed_line, styles['Normal']))
            
            doc.build(story)
        
        else:
            raise ImportError("Neither weasyprint nor reportlab is available")
    
    def export_as_rtf(self, file_path):
        """Export as Rich Text Format."""
        markdown_text = self.editor.toPlainText()
        html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
        
        # Simple RTF header
        rtf_content = r"{\rtf1\ansi\deff0"
        rtf_content += r"{\fonttbl{\f0 Times New Roman;}}"
        rtf_content += r"{\colortbl;\red0\green0\blue0;}"
        rtf_content += r"\fs24"
        
        # Parse and add content
        lines = markdown_text.split('\n')
        in_code_block = False
        code_content = []
        
        for line in lines:
            if line.strip().startswith('```'):
                if in_code_block:
                    # End of code block
                    if code_content:
                        rtf_content += r"{\pard\plain\f0\fs20 "
                        for code_line in code_content:
                            rtf_content += code_line.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"\line "
                        rtf_content += r"}\par"
                        code_content = []
                    in_code_block = False
                else:
                    # Start of code block
                    in_code_block = True
                continue
            
            if in_code_block:
                code_content.append(line)
                continue
            
            # Headers
            if line.startswith('# '):
                rtf_content += r"{\pard\plain\f0\fs36\b " + line[2:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.startswith('## '):
                rtf_content += r"{\pard\plain\f0\fs32\b " + line[3:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.startswith('### '):
                rtf_content += r"{\pard\plain\f0\fs28\b " + line[4:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.startswith('#### '):
                rtf_content += r"{\pard\plain\f0\fs26\b " + line[5:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.startswith('##### '):
                rtf_content += r"{\pard\plain\f0\fs24\b " + line[6:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.startswith('###### '):
                rtf_content += r"{\pard\plain\f0\fs22\b " + line[7:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            # Horizontal rule
            elif line.strip() == '---' or line.strip() == '***':
                rtf_content += r"{\pard\plain\f0\fs24 " + "_" * 50 + r"}\par"
            # Lists
            elif line.strip().startswith(('- ', '* ', '+ ')):
                rtf_content += r"{\pard\plain\f0\fs24 \bullet " + line.strip()[2:].replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            elif line.strip() and line.strip()[0].isdigit() and line.strip().find('.') == 1:
                rtf_content += r"{\pard\plain\f0\fs24 " + line.strip().replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
            # Empty line
            elif not line.strip():
                rtf_content += r"\par"
            # Regular paragraph
            elif line.strip():
                # Simple markdown processing
                processed_line = line
                # Bold
                processed_line = re.sub(r'\*\*(.*?)\*\*', r'\b \1\b0', processed_line)
                # Italic
                processed_line = re.sub(r'\*(.*?)\*', r'\i \1\i0', processed_line)
                # Code
                processed_line = re.sub(r'`(.*?)`', r'\f1\fs18 \1\f0\fs24', processed_line)
                # Links - simplified for RTF
                processed_line = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', processed_line)
                
                rtf_content += r"{\pard\plain\f0\fs24 " + processed_line.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}') + r"}\par"
        
        rtf_content += "}"
        
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(rtf_content)
    
    def export_as_odt(self, file_path):
        """Export as OpenDocument Text (ODT)."""
        # ODT is essentially a ZIP file with XML content
        # For simplicity, we'll create a basic ODT structure
        import zipfile
        import xml.etree.ElementTree as ET
        
        markdown_text = self.editor.toPlainText()
        html = markdown.markdown(markdown_text, extensions=['codehilite', 'tables', 'toc'])
        
        # Create content.xml
        content = ET.Element("office:document-content")
        content.set("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        content.set("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        content.set("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        content.set("office:version", "1.0")
        
        body = ET.SubElement(content, "office:body")
        text = ET.SubElement(body, "office:text")
        
        # Parse and add content
        lines = markdown_text.split('\n')
        in_code_block = False
        code_content = []
        
        for line in lines:
            if line.strip().startswith('```'):
                if in_code_block:
                    # End of code block
                    if code_content:
                        for code_line in code_content:
                            p = ET.SubElement(text, "text:p")
                            p.text = code_line
                        code_content = []
                    in_code_block = False
                else:
                    # Start of code block
                    in_code_block = True
                continue
            
            if in_code_block:
                code_content.append(line)
                continue
            
            # Headers
            if line.startswith('# '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "1"})
                h.text = line[2:]
            elif line.startswith('## '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "2"})
                h.text = line[3:]
            elif line.startswith('### '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "3"})
                h.text = line[4:]
            elif line.startswith('#### '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "4"})
                h.text = line[5:]
            elif line.startswith('##### '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "5"})
                h.text = line[6:]
            elif line.startswith('###### '):
                h = ET.SubElement(text, "text:h", attrib={"text:outline-level": "6"})
                h.text = line[7:]
            # Horizontal rule
            elif line.strip() == '---' or line.strip() == '***':
                p = ET.SubElement(text, "text:p")
                p.text = "_" * 50
            # Lists
            elif line.strip().startswith(('- ', '* ', '+ ')):
                p = ET.SubElement(text, "text:p")
                p.text = "• " + line.strip()[2:]
            elif line.strip() and line.strip()[0].isdigit() and line.strip().find('.') == 1:
                p = ET.SubElement(text, "text:p")
                p.text = line.strip()
            # Empty line
            elif not line.strip():
                p = ET.SubElement(text, "text:p")
                p.text = ""
            # Regular paragraph
            elif line.strip():
                # Simple markdown processing
                processed_line = line
                # Bold - simplified for ODT
                processed_line = re.sub(r'\*\*(.*?)\*\*', r'\1', processed_line)
                # Italic
                processed_line = re.sub(r'\*(.*?)\*', r'\1', processed_line)
                # Code
                processed_line = re.sub(r'`(.*?)`', r'\1', processed_line)
                # Links
                processed_line = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', processed_line)
                
                p = ET.SubElement(text, "text:p")
                p.text = processed_line
        
        # Create minimal ODT structure
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as odt:
            # mimetype
            odt.writestr('mimetype', 'application/vnd.oasis.opendocument.text')
            
            # content.xml
            content_str = ET.tostring(content, encoding='unicode', xml_declaration=True)
            odt.writestr('content.xml', content_str)
            
            # Basic manifest
            manifest = ET.Element("manifest:manifest")
            manifest.set("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
            
            file_entry1 = ET.SubElement(manifest, "manifest:file-entry")
            file_entry1.set("manifest:full-path", "/")
            file_entry1.set("manifest:media-type", "application/vnd.oasis.opendocument.text")
            
            file_entry2 = ET.SubElement(manifest, "manifest:file-entry")
            file_entry2.set("manifest:full-path", "content.xml")
            file_entry2.set("manifest:media-type", "text/xml")
            
            manifest_str = ET.tostring(manifest, encoding='unicode', xml_declaration=True)
            odt.writestr('META-INF/manifest.xml', manifest_str)


def main():
    app = QApplication(sys.argv)
    editor = MarkdownEditor()
    editor.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
