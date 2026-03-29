"""Find / Replace dialog for Smart Markdown Editor.

This module provides a standalone, reusable dialog that operates on any
QTextEdit passed at construction time.  The dialog is non-modal so the
user can keep editing while searching.
"""

from __future__ import annotations

from PySide6.QtGui import QTextCursor, QTextDocument
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)


class FindReplaceDialog(QDialog):
    """Non-modal Find / Replace dialog bound to a specific QTextEdit."""

    def __init__(self, editor: QTextEdit, parent=None) -> None:
        super().__init__(parent)
        self._editor = editor
        self.setWindowTitle("Find / Replace")
        self.setModal(False)
        self._build_ui()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def open_find_mode(self) -> None:
        """Show the dialog with only Find controls enabled."""
        self._replace_input.setEnabled(False)
        self._replace_btn.setEnabled(False)
        self._replace_all_btn.setEnabled(False)
        self._show_and_focus()

    def open_replace_mode(self) -> None:
        """Show the dialog with all Find & Replace controls enabled."""
        self._replace_input.setEnabled(True)
        self._replace_btn.setEnabled(True)
        self._replace_all_btn.setEnabled(True)
        self._show_and_focus()

    def apply_theme(self, dark_mode: bool) -> None:
        """Update the dialog stylesheet for the current theme."""
        from src.ui.themes import ThemeManager  # local import avoids circular dep
        self.setStyleSheet(ThemeManager.get_dialog_stylesheet(dark_mode))

    # ------------------------------------------------------------------
    # Find / Replace operations
    # ------------------------------------------------------------------

    def find_text(self, backward: bool = False) -> bool:
        """Find the next (or previous) occurrence of the search term.

        Returns True when a match is found.
        """
        needle = self._find_input.text()
        if not needle:
            return False

        doc = self._editor.document()
        cursor = self._editor.textCursor()
        flags = self._build_find_flags(backward)

        found = doc.find(needle, cursor, flags)
        if found.isNull():
            # Wrap around
            wrap = self._editor.textCursor()
            if backward:
                wrap.movePosition(QTextCursor.MoveOperation.End)
            else:
                wrap.movePosition(QTextCursor.MoveOperation.Start)
            found = doc.find(needle, wrap, flags)

        if found.isNull():
            QMessageBox.information(self, "Find", f"'{needle}' not found.")
            return False

        self._editor.setTextCursor(found)
        self._editor.ensureCursorVisible()
        return True

    def replace_one(self) -> None:
        """Replace the current selection if it matches, then find the next."""
        needle = self._find_input.text()
        if not needle:
            return

        replacement = self._replace_input.text()
        cursor = self._editor.textCursor()

        if cursor.hasSelection():
            selected = cursor.selectedText()
            match_case = self._match_case_cb.isChecked()
            matches = selected == needle if match_case else selected.lower() == needle.lower()
        else:
            matches = False

        if not matches:
            if not self.find_text(backward=False):
                return
            cursor = self._editor.textCursor()

        cursor.beginEditBlock()
        cursor.insertText(replacement)
        cursor.endEditBlock()
        self.find_text(backward=False)

    def replace_all(self) -> None:
        """Replace all occurrences of the search term."""
        needle = self._find_input.text()
        if not needle:
            return

        replacement = self._replace_input.text()
        doc = self._editor.document()

        flags = QTextDocument.FindFlag(0)
        if self._match_case_cb.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively

        outer_cursor = self._editor.textCursor()
        outer_cursor.beginEditBlock()

        scan = self._editor.textCursor()
        scan.movePosition(QTextCursor.MoveOperation.Start)
        count = 0
        while True:
            found = doc.find(needle, scan, flags)
            if found.isNull():
                break
            found.insertText(replacement)
            count += 1
            scan = found

        outer_cursor.endEditBlock()
        QMessageBox.information(self, "Replace All", f"Replaced {count} occurrence(s).")

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        layout = QVBoxLayout()

        # Find row
        find_row = QHBoxLayout()
        find_row.addWidget(QLabel("Find:"))
        self._find_input = QLineEdit()
        self._find_input.setPlaceholderText("Text to find")
        find_row.addWidget(self._find_input)
        layout.addLayout(find_row)

        # Replace row
        replace_row = QHBoxLayout()
        replace_row.addWidget(QLabel("Replace:"))
        self._replace_input = QLineEdit()
        self._replace_input.setPlaceholderText("Replacement text")
        replace_row.addWidget(self._replace_input)
        layout.addLayout(replace_row)

        # Options row
        options_row = QHBoxLayout()
        self._match_case_cb = QCheckBox("Match case")
        options_row.addWidget(self._match_case_cb)
        options_row.addStretch()
        layout.addLayout(options_row)

        # Buttons row
        buttons_row = QHBoxLayout()
        self._find_prev_btn = QPushButton("Find Previous")
        self._find_next_btn = QPushButton("Find Next")
        self._replace_btn = QPushButton("Replace")
        self._replace_all_btn = QPushButton("Replace All")
        close_btn = QPushButton("Close")

        self._find_prev_btn.clicked.connect(lambda: self.find_text(backward=True))
        self._find_next_btn.clicked.connect(lambda: self.find_text(backward=False))
        self._replace_btn.clicked.connect(self.replace_one)
        self._replace_all_btn.clicked.connect(self.replace_all)
        close_btn.clicked.connect(self.close)

        for btn in (
            self._find_prev_btn,
            self._find_next_btn,
            self._replace_btn,
            self._replace_all_btn,
            close_btn,
        ):
            buttons_row.addWidget(btn)
        layout.addLayout(buttons_row)

        self.setLayout(layout)

        # Keyboard shortcuts
        self._find_input.returnPressed.connect(lambda: self.find_text(backward=False))
        self._replace_input.returnPressed.connect(self.replace_one)

    def _build_find_flags(self, backward: bool) -> QTextDocument.FindFlag:
        flags = QTextDocument.FindFlag(0)
        if backward:
            flags |= QTextDocument.FindFlag.FindBackward
        if self._match_case_cb.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively
        return flags

    def _show_and_focus(self) -> None:
        self.show()
        self.raise_()
        self.activateWindow()
        self._find_input.setFocus()
        self._find_input.selectAll()
