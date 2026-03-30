"""Main application window for Smart Markdown Editor.

This module contains :class:`MainWindow`, the top-level QMainWindow that
wires together all sub-modules:

* :class:`src.core.analyzer.MarkdownAnalyzer`  — document metrics
* :class:`src.core.highlighter.MarkdownSyntaxHighlighter` — editor colours
* :class:`src.ui.assistant_panel.AssistantPanel` — side-panel widget
* :class:`src.ui.dialogs.FindReplaceDialog` — find / replace
* :class:`src.ui.themes.ThemeManager` — stylesheet / HTML builder
* :mod:`src.exporters` — exporter registry
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path

import markdown
from PySide6.QtCore import Qt, QSettings, QTimer
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import (
    QFileDialog,
    QMainWindow,
    QMenu,
    QMessageBox,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.config import (
    APP_NAME,
    APP_ORGANIZATION,
    ANALYSIS_UPDATE_DELAY_MS,
    DEFAULT_AUTOSAVE_INTERVAL_MS,
    MAX_FILE_SIZE_BYTES,
    PREVIEW_UPDATE_DELAY_MS,
)
from src.core.analyzer import MarkdownAnalyzer
from src.core.highlighter import MarkdownSyntaxHighlighter
from src.exporters import get_available_exporters, get_exporter
from src.ui.assistant_panel import AssistantPanel
from src.ui.dialogs import FindReplaceDialog
from src.ui.themes import ThemeManager

try:
    from pygments.formatters import HtmlFormatter

    PYGMENTS_AVAILABLE = True
except ImportError:
    PYGMENTS_AVAILABLE = False

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """Top-level application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setGeometry(100, 100, 1400, 900)

        self._settings = QSettings(APP_ORGANIZATION, APP_NAME)
        self._dark_mode: bool = bool(
            self._settings.value("darkMode", False, type=bool)
        )
        self.current_file: str | None = None
        self._custom_preview_css_path: str = str(
            self._settings.value("previewCssPath", "", type=str)
        )
        self._custom_preview_css_cache: str = ""
        self._custom_preview_css_cache_mtime: float | None = None
        self._recent_files: list[str] = self._load_recent_files()
        self._pygments_css_by_theme: dict[str, str | None] = {
            "light": None,
            "dark": None,
        }

        self._build_ui()
        self._build_timers()

        self.update_preview()
        self.update_analysis()

    # ==========================================================================
    # UI construction
    # ==========================================================================

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Editor
        self.editor = QTextEdit()
        self.editor.setPlaceholderText("Type your markdown here…")
        self.editor.textChanged.connect(self._on_text_changed)
        self._highlighter = MarkdownSyntaxHighlighter(
            self.editor.document(), dark_mode=self._dark_mode
        )

        # Preview
        self.preview = QWebEngineView()

        # Assistant panel
        self.assistant_panel = AssistantPanel()
        self.assistant_panel.format_button.clicked.connect(self.auto_format_document)

        splitter.addWidget(self.editor)
        splitter.addWidget(self.preview)
        splitter.addWidget(self.assistant_panel)
        splitter.setSizes([420, 630, 350])

        layout = QVBoxLayout()
        layout.addWidget(splitter)
        central.setLayout(layout)

        self._build_menu_bar()
        self._apply_theme()

        # Find/Replace — created lazily on first use
        self._find_replace_dialog: FindReplaceDialog | None = None

    def _build_timers(self) -> None:
        self._preview_timer = QTimer(singleShot=True)
        self._preview_timer.timeout.connect(self.update_preview)

        self._analysis_timer = QTimer(singleShot=True)
        self._analysis_timer.timeout.connect(self.update_analysis)

        interval = int(  # type: ignore[call-overload]
            self._settings.value(
                "autoSaveIntervalMs", DEFAULT_AUTOSAVE_INTERVAL_MS, type=int
            )
        )
        self._autosave_timer = QTimer()
        self._autosave_timer.setInterval(interval)
        self._autosave_timer.timeout.connect(self._autosave_tick)
        self._autosave_timer.start()

    def _build_menu_bar(self) -> None:
        menubar = self.menuBar()

        # ---- File ----
        file_menu = menubar.addMenu("File")

        new_action = file_menu.addAction("New")
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_file)

        open_action = file_menu.addAction("Open")
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)

        self._recent_files_menu = file_menu.addMenu("Recent Files")
        self._rebuild_recent_files_menu()

        save_action = file_menu.addAction("Save")
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)

        save_as_action = file_menu.addAction("Save As…")
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_file_as)

        file_menu.addSeparator()

        export_menu = file_menu.addMenu("Export As")
        self._populate_export_menu(export_menu)

        file_menu.addSeparator()

        exit_action = file_menu.addAction("Exit")
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)

        # ---- Edit ----
        edit_menu = menubar.addMenu("Edit")

        undo_action = edit_menu.addAction("Undo")
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.editor.undo)

        redo_action = edit_menu.addAction("Redo")
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.editor.redo)

        edit_menu.addSeparator()

        find_action = edit_menu.addAction("Find…")
        find_action.setShortcut("Ctrl+F")
        find_action.triggered.connect(self.open_find_dialog)

        replace_action = edit_menu.addAction("Replace…")
        replace_action.setShortcut("Ctrl+H")
        replace_action.triggered.connect(self.open_replace_dialog)

        # ---- View ----
        view_menu = menubar.addMenu("View")

        self._dark_mode_action = view_menu.addAction("Dark Mode")
        self._dark_mode_action.setCheckable(True)
        self._dark_mode_action.setChecked(self._dark_mode)
        self._dark_mode_action.triggered.connect(self._toggle_dark_mode)

        view_menu.addSeparator()

        preview_css_action = view_menu.addAction("Preview CSS…")
        preview_css_action.triggered.connect(self.choose_preview_css)

        clear_css_action = view_menu.addAction("Clear Preview CSS")
        clear_css_action.triggered.connect(self.clear_preview_css)

    def _populate_export_menu(self, menu: QMenu) -> None:
        """Add one action per available exporter to *menu*."""
        for exporter_cls in get_available_exporters():
            exporter = exporter_cls()
            action = menu.addAction(f"{exporter.name} (.{exporter.extension})")
            action.triggered.connect(
                lambda checked=False, ext=exporter.extension: self.export_file(ext)
            )

    # ==========================================================================
    # Theme
    # ==========================================================================

    def _apply_theme(self) -> None:
        self.editor.setStyleSheet(
            ThemeManager.get_editor_stylesheet(self._dark_mode)
        )
        self.assistant_panel.setStyleSheet(
            ThemeManager.get_panel_stylesheet(self._dark_mode)
        )
        if self._highlighter is not None:
            self._highlighter.set_dark_mode(self._dark_mode)
        if self._find_replace_dialog is not None:
            self._find_replace_dialog.apply_theme(self._dark_mode)

    def _toggle_dark_mode(self, checked: bool) -> None:
        self._dark_mode = bool(checked)
        self._settings.setValue("darkMode", self._dark_mode)
        self._apply_theme()
        self.update_preview()

    # ==========================================================================
    # Preview
    # ==========================================================================

    def update_preview(self) -> None:
        markdown_text = self.editor.toPlainText()
        html_body = markdown.markdown(
            markdown_text, extensions=["codehilite", "tables", "toc"]
        )

        theme_key = "dark" if self._dark_mode else "light"
        pygments_css = ""
        if PYGMENTS_AVAILABLE:
            if self._pygments_css_by_theme[theme_key] is None:
                style_name = "monokai" if self._dark_mode else "default"
                self._pygments_css_by_theme[theme_key] = (
                    HtmlFormatter(style=style_name).get_style_defs(".codehilite")
                )
            pygments_css = self._pygments_css_by_theme[theme_key] or ""

        custom_css = self._get_custom_preview_css()
        html = ThemeManager.build_preview_html(
            html_body,
            dark_mode=self._dark_mode,
            custom_css=custom_css,
            pygments_css=pygments_css,
        )
        self.preview.setHtml(html)

    # ==========================================================================
    # Analysis
    # ==========================================================================

    def update_analysis(self) -> None:
        markdown_text = self.editor.toPlainText()
        if not markdown_text.strip():
            self.assistant_panel.clear()
            return
        metrics = MarkdownAnalyzer(markdown_text).analyze()
        self.assistant_panel.update_metrics(metrics)

    # ==========================================================================
    # Find / Replace
    # ==========================================================================

    def _get_find_replace_dialog(self) -> FindReplaceDialog:
        if self._find_replace_dialog is None:
            self._find_replace_dialog = FindReplaceDialog(self.editor, parent=self)
            self._find_replace_dialog.apply_theme(self._dark_mode)
        return self._find_replace_dialog

    def open_find_dialog(self) -> None:
        self._get_find_replace_dialog().open_find_mode()

    def open_replace_dialog(self) -> None:
        self._get_find_replace_dialog().open_replace_mode()

    # ==========================================================================
    # File operations
    # ==========================================================================

    def new_file(self) -> None:
        self.editor.clear()
        self.current_file = None
        self.editor.document().setModified(False)

    def open_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Markdown File",
            "",
            "Markdown Files (*.md);;All Files (*)",
        )
        if file_path:
            self.open_file_path(file_path)

    def open_file_path(self, file_path: str) -> None:
        if not file_path:
            return
        try:
            try:
                size = os.path.getsize(file_path)
                if size > MAX_FILE_SIZE_BYTES:
                    QMessageBox.warning(
                        self,
                        "File Too Large",
                        f"Cannot open file: it exceeds the 10 MB limit "
                        f"({size / (1024 * 1024):.1f} MB).",
                    )
                    return
            except OSError:
                pass
            with open(file_path, "r", encoding="utf-8") as fh:
                content = fh.read()
            self.editor.setPlainText(content)
            self.editor.document().setModified(False)
            self.current_file = file_path
            self._add_recent_file(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Could not open file: {exc}")

    def save_file(self) -> None:
        if self.current_file:
            self._save_to_file(self.current_file, show_errors=True, update_recent=True)
        else:
            self.save_file_as()

    def save_file_as(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Markdown File",
            "",
            "Markdown Files (*.md);;All Files (*)",
        )
        if file_path:
            self._save_to_file(file_path, show_errors=True, update_recent=True)

    def _save_to_file(
        self, file_path: str, *, show_errors: bool, update_recent: bool
    ) -> bool:
        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                fh.write(self.editor.toPlainText())
            self.current_file = file_path
            self.editor.document().setModified(False)
            if update_recent:
                self._add_recent_file(file_path)
            return True
        except Exception as exc:
            if show_errors:
                QMessageBox.critical(self, "Error", f"Could not save file: {exc}")
            return False

    def _autosave_tick(self) -> None:
        if not self.current_file:
            return
        if not self.editor.document().isModified():
            return
        self._save_to_file(self.current_file, show_errors=False, update_recent=False)

    # ==========================================================================
    # Custom preview CSS
    # ==========================================================================

    def choose_preview_css(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Preview CSS", "", "CSS Files (*.css);;All Files (*)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".css"):
            QMessageBox.warning(
                self,
                "Invalid File",
                "Please select a valid CSS file (.css extension required).",
            )
            return
        self._custom_preview_css_path = file_path
        self._custom_preview_css_cache = ""
        self._custom_preview_css_cache_mtime = None
        self._settings.setValue("previewCssPath", file_path)
        self.update_preview()

    def clear_preview_css(self) -> None:
        self._custom_preview_css_path = ""
        self._custom_preview_css_cache = ""
        self._custom_preview_css_cache_mtime = None
        self._settings.setValue("previewCssPath", "")
        self.update_preview()

    def _get_custom_preview_css(self) -> str:
        path = (self._custom_preview_css_path or "").strip()
        if not path:
            return ""
        if not os.path.exists(path):
            self.clear_preview_css()
            return ""
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return ""
        if self._custom_preview_css_cache_mtime != mtime:
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    self._custom_preview_css_cache = fh.read()
                self._custom_preview_css_cache_mtime = mtime
            except Exception:
                return ""
        return self._custom_preview_css_cache

    # ==========================================================================
    # Recent files
    # ==========================================================================

    def _load_recent_files(self) -> list[str]:
        value = self._settings.value("recentFiles", [])
        if value is None:
            return []
        if isinstance(value, str):
            return [value]
        if isinstance(value, (list, tuple)):
            return [str(v) for v in value if v]
        return []

    def _save_recent_files(self) -> None:
        self._settings.setValue("recentFiles", self._recent_files)

    def _add_recent_file(self, file_path: str) -> None:
        if not file_path:
            return
        file_path = os.path.abspath(file_path)
        self._recent_files = [
            p for p in self._recent_files if os.path.abspath(p) != file_path
        ]
        self._recent_files.insert(0, file_path)
        self._recent_files = self._recent_files[:10]
        self._save_recent_files()
        self._rebuild_recent_files_menu()

    def _clear_recent_files(self) -> None:
        self._recent_files = []
        self._save_recent_files()
        self._rebuild_recent_files_menu()

    def _rebuild_recent_files_menu(self) -> None:
        if not hasattr(self, "_recent_files_menu") or self._recent_files_menu is None:
            return
        self._recent_files_menu.clear()

        existing = [p for p in self._recent_files if p and os.path.exists(p)]
        self._recent_files = existing
        self._save_recent_files()

        if not self._recent_files:
            empty_action = self._recent_files_menu.addAction("(No recent files)")
            empty_action.setEnabled(False)
        else:
            for path in self._recent_files:
                action = self._recent_files_menu.addAction(path)
                action.triggered.connect(
                    lambda checked=False, p=path: self.open_file_path(p)
                )

        self._recent_files_menu.addSeparator()
        clear_action = self._recent_files_menu.addAction("Clear Recent Files")
        clear_action.setEnabled(bool(self._recent_files))
        clear_action.triggered.connect(self._clear_recent_files)

    # ==========================================================================
    # Export
    # ==========================================================================

    def export_file(self, extension: str) -> None:
        exporter_cls = get_exporter(extension)
        if exporter_cls is None:
            QMessageBox.warning(
                self, "Export Error", f"No exporter registered for '.{extension}'."
            )
            return

        exporter = exporter_cls()
        if not exporter.is_available:
            QMessageBox.warning(
                self,
                "Export Error",
                f"The exporter for '.{extension}' requires an additional library.\n"
                f"{exporter.get_install_message()}",
            )
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Export as {extension.upper()}",
            "",
            exporter.file_filter,
        )
        if not file_path:
            return

        try:
            exporter.export(self.editor.toPlainText(), Path(file_path))
            QMessageBox.information(
                self, "Export Successful", f"Exported to:\n{file_path}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Error", f"Export failed: {exc}")

    # ==========================================================================
    # Auto-format
    # ==========================================================================

    def auto_format_document(self) -> None:
        text = self.editor.toPlainText()
        if not text.strip():
            QMessageBox.information(
                self, "Auto-Format", "Document is empty. Nothing to format."
            )
            return
        self.editor.setPlainText(self._format_markdown(text))
        QMessageBox.information(self, "Auto-Format", "Document formatted successfully!")

    @staticmethod
    def _format_markdown(text: str) -> str:
        """Apply auto-formatting rules to markdown text."""
        lines = text.split("\n")
        formatted: list[str] = []
        prev_was_empty = False

        for i, line in enumerate(lines):
            stripped = line.strip()

            # Insert blank line before headings
            if stripped.startswith("#") and i > 0 and not prev_was_empty:
                formatted.append("")

            # Ensure space after heading marker
            if stripped.startswith("#"):
                m = re.match(r"^(#{1,6})(\S)", stripped)
                if m:
                    level = m.group(1)
                    rest = stripped[len(level):]
                    formatted.append(f"{level} {rest}")
                    prev_was_empty = False
                    continue

            # Fix list markers
            if re.match(r"^(\s*)([-*+])(\S)", line):
                m = re.match(r"^(\s*)([-*+])(.*)$", line)
                if m:
                    formatted.append(
                        f"{m.group(1)}{m.group(2)} {m.group(3).strip()}"
                    )
                    prev_was_empty = False
                    continue

            if re.match(r"^(\s*)(\d+\.)(\S)", line):
                m = re.match(r"^(\s*)(\d+\.)(.*)$", line)
                if m:
                    formatted.append(
                        f"{m.group(1)}{m.group(2)} {m.group(3).strip()}"
                    )
                    prev_was_empty = False
                    continue

            if not stripped:
                if not prev_was_empty:
                    formatted.append("")
                    prev_was_empty = True
            else:
                formatted.append(line)
                prev_was_empty = False

        # Strip trailing blank lines
        while formatted and not formatted[-1]:
            formatted.pop()

        return "\n".join(formatted)

    # ==========================================================================
    # Internal slots
    # ==========================================================================

    def _on_text_changed(self) -> None:
        self._preview_timer.start(PREVIEW_UPDATE_DELAY_MS)
        self._analysis_timer.start(ANALYSIS_UPDATE_DELAY_MS)
