# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Added

- `src/ui/` package with four dedicated modules:
  - `themes.py` — `ThemeManager` with static methods for editor, panel, dialog
    and preview-HTML stylesheets; removes all raw CSS from business-logic code.
  - `dialogs.py` — `FindReplaceDialog(QDialog)`: a self-contained, non-modal
    Find / Replace dialog that operates on any `QTextEdit` instance.
  - `assistant_panel.py` — `AssistantPanel(QWidget)`: the Smart Assistant
    side-panel with `update_metrics()` and `clear()` public API.
  - `main_window.py` — `MainWindow(QMainWindow)`: a clean, fully-modular
    replacement for the `MarkdownEditor` monolith; delegates styling to
    `ThemeManager`, analysis to `MarkdownAnalyzer`, and exports to the registry.
- `src/exporters/builtin.py` — seven concrete `BaseExporter` subclasses
  (`MarkdownExporter`, `TextExporter`, `HTMLExporter`, `DocxExporter`,
  `PDFExporter`, `RTFExporter`, `ODTExporter`) and a `register_all()` function.
- `src/exporters/__init__.py` now exposes a clean registry API; `get_available_exporters()`
  correctly instantiates each exporter class before checking `is_available`.
- `CONTRIBUTING.md` — developer setup, code style guide, and PR workflow.
- `CHANGELOG.md` — this file.

### Changed

- `src/main.py` — complete rewrite: clean `setup_logging()` and `main()`
  entry-point; replaced fragmented one-symbol imports and fixed syntax errors.
- `src/core/analyzer.py`:
  - `MarkdownAnalyzer.__init__` now accepts `None` and non-string inputs
    gracefully (converts to empty string) rather than crashing.
  - `_analyze_headings()` and `_count_words()` are now memoised with instance
    caches to avoid redundant computation on repeated calls.
- `src/config.py` — trailing XML artifact removed.
- `src/core/highlighter.py` — trailing XML artifact removed.
- `src/utils/__init__.py` — trailing XML artifact removed.
- `tests/test_analyzer.py`:
  - `test_none_input` and `test_numeric_input` corrected to match new guarded
    behaviour.
  - Added `test_unicode_content`, `test_only_code_blocks`, and
    `test_headings_cache_consistency`.
- `test_exports.py` — converted from `print()`-based script to a proper
  `pytest` suite with assertions, `tmp_path` fixture, and conditional skipping
  for optional library exporters.
- `.github/workflows/ci.yml`:
  - Added Python version matrix (`3.9`, `3.12`).
  - Added `apt-get` step for WeasyPrint system libraries (Pango, Cairo, …).
  - Tests and build now run under `xvfb-run` for headless PySide6 support.
  - Removed invalid `--ignore src/legacy/` flag from Ruff command.
  - `test_exports.py` included in test run.
  - Fixed build job to reference the correct entry-point.

### Fixed

- `markdown_editor.py`:
  - Readability scoring logic was inverted: the > 150 word penalty (−20) was
    listed after the > 100 penalty (−10), so only the weaker penalty ever fired.
    Order is now correct.
  - `open_file_path()`: added a 10 MB file-size guard before reading the file
    into memory.
  - `choose_preview_css()`: rejects files without a `.css` extension to prevent
    loading arbitrary file content.
  - `export_as_text()`: removed duplicate `import re` that shadowed the
    module-level import.
  - `export_as_docx()`, `export_as_rtf()`, `export_as_pdf()`, `export_as_odt()`:
    all reformatted to use the new shared `_parse_markdown_lines()` iterator,
    eliminating four copies of the same hand-written line-by-line parser and the
    associated dead/duplicate code.
  - `export_as_odt()`: removed the duplicate XML namespace setup and second
    `zipfile.ZipFile` write block that were left behind by an earlier incomplete
    refactor.
  - `format_markdown()`: removed stale `prev_was_heading` variable assignments
    that were never read.
  - Added shared `_parse_markdown_lines()` and `_strip_inline_md()` helper methods.

---

## [1.0.0]

### Added

- Initial release of Smart Markdown Editor.
- Live markdown preview via `QWebEngineView`.
- Syntax highlighting in the editor (`MarkdownSyntaxHighlighter`).
- Smart Assistant panel with document statistics and quality analysis.
- Export to Markdown, plain text, HTML, DOCX, PDF, RTF, and ODT.
- Find & Replace dialog.
- Recent files menu with auto-cleanup of missing paths.
- Auto-save (configurable interval, defaults to 30 s).
- Custom preview CSS support.
