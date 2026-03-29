"""Theme management for Smart Markdown Editor.

All theme-related stylesheet strings and the preview HTML template are
centralised here so the rest of the codebase stays free of raw CSS.
"""

from __future__ import annotations

from src.config import LightTheme, DarkTheme


class ThemeManager:
    """Provides QSS stylesheets and HTML templates for both light and dark modes."""

    # ------------------------------------------------------------------
    # Editor (QTextEdit)
    # ------------------------------------------------------------------

    @staticmethod
    def get_editor_stylesheet(dark_mode: bool) -> str:
        """Return QSS for the main text editor widget."""
        if dark_mode:
            t = DarkTheme
            return (
                f"QTextEdit {{"
                f"  background-color: {t.EDITOR_BG};"
                f"  color: {t.EDITOR_FG};"
                f"  border: 1px solid {t.EDITOR_BORDER};"
                f"  font-family: 'SF Mono', 'Monaco', 'Menlo', 'Consolas', monospace;"
                f"  font-size: 14px;"
                f"  padding: 10px;"
                f"  selection-background-color: {t.SELECTION_BG};"
                f"  selection-color: {t.SELECTION_FG};"
                f"}}"
            )
        t = LightTheme  # type: ignore[assignment]
        return (
            f"QTextEdit {{"
            f"  background-color: {t.EDITOR_BG};"
            f"  color: {t.EDITOR_FG};"
            f"  border: 1px solid {t.EDITOR_BORDER};"
            f"  font-family: 'SF Mono', 'Monaco', 'Menlo', 'Consolas', monospace;"
            f"  font-size: 14px;"
            f"  padding: 10px;"
            f"  selection-background-color: {t.SELECTION_BG};"
            f"  selection-color: {t.SELECTION_FG};"
            f"}}"
        )

    # ------------------------------------------------------------------
    # Assistant / side panels
    # ------------------------------------------------------------------

    @staticmethod
    def get_panel_stylesheet(dark_mode: bool) -> str:
        """Return QSS for the assistant / sidebar panel widget."""
        if dark_mode:
            t = DarkTheme
            return (
                f"QWidget {{"
                f"  background-color: {t.EDITOR_BG};"
                f"  color: {t.EDITOR_FG};"
                f"}}"
                f"QGroupBox {{"
                f"  border: 1px solid {t.EDITOR_BORDER};"
                f"  margin-top: 8px;"
                f"  padding: 8px;"
                f"}}"
                f"QGroupBox::title {{"
                f"  subcontrol-origin: margin;"
                f"  left: 10px;"
                f"  padding: 0 4px 0 4px;"
                f"}}"
                f"QPushButton {{"
                f"  border: 1px solid {t.EDITOR_BORDER};"
                f"  padding: 6px 10px;"
                f"}}"
            )
        return ""

    # ------------------------------------------------------------------
    # Dialogs
    # ------------------------------------------------------------------

    @staticmethod
    def get_dialog_stylesheet(dark_mode: bool) -> str:
        """Return QSS for modal/non-modal dialogs."""
        if dark_mode:
            t = DarkTheme
            return (
                f"QDialog {{ background-color: {t.EDITOR_BG}; color: {t.EDITOR_FG}; }}"
                f"QLabel {{ color: {t.EDITOR_FG}; }}"
                f"QLineEdit {{"
                f"  background-color: {t.CODE_BG};"
                f"  color: {t.EDITOR_FG};"
                f"  border: 1px solid {t.EDITOR_BORDER};"
                f"  padding: 4px;"
                f"}}"
                f"QCheckBox {{ color: {t.EDITOR_FG}; }}"
                f"QPushButton {{ border: 1px solid {t.EDITOR_BORDER}; padding: 6px 10px; }}"
            )
        return ""

    # ------------------------------------------------------------------
    # Preview HTML
    # ------------------------------------------------------------------

    @staticmethod
    def build_preview_html(
        html_body: str,
        *,
        dark_mode: bool,
        custom_css: str = "",
        pygments_css: str = "",
    ) -> str:
        """Wrap *html_body* in a fully styled HTML document for QWebEngineView."""
        if dark_mode:
            body_bg = DarkTheme.EDITOR_BG
            body_fg = DarkTheme.EDITOR_FG
            border = DarkTheme.EDITOR_BORDER
            muted = DarkTheme.MUTED_COLOR
            link = DarkTheme.LINK_COLOR
            code_bg = DarkTheme.CODE_BG
        else:
            body_bg = "#fff"
            body_fg = "#333"
            border = "#eaecef"
            muted = LightTheme.MUTED_COLOR
            link = "#0366d6"
            code_bg = LightTheme.CODE_BG

        return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    {pygments_css}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      line-height: 1.6;
      color: {body_fg};
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
      background-color: {body_bg};
    }}
    h1, h2, h3, h4, h5, h6 {{
      margin-top: 24px;
      margin-bottom: 16px;
      font-weight: 600;
      line-height: 1.25;
    }}
    h1 {{ font-size: 2em; border-bottom: 1px solid {border}; padding-bottom: 0.3em; }}
    h2 {{ font-size: 1.5em; border-bottom: 1px solid {border}; padding-bottom: 0.3em; }}
    h3 {{ font-size: 1.25em; }}
    h4 {{ font-size: 1em; }}
    h5 {{ font-size: 0.875em; }}
    h6 {{ font-size: 0.85em; color: {muted}; }}
    p {{ margin-bottom: 16px; }}
    code {{
      background-color: {code_bg};
      border-radius: 3px;
      font-size: 85%;
      padding: 0.2em 0.4em;
    }}
    pre {{
      background-color: {code_bg};
      border-radius: 6px;
      padding: 16px;
      overflow: auto;
      font-size: 85%;
      line-height: 1.45;
    }}
    .codehilite {{
      background-color: {code_bg};
      border-radius: 6px;
      padding: 16px;
      overflow: auto;
      margin-bottom: 16px;
    }}
    .codehilite pre {{ margin: 0; padding: 0; background: transparent; }}
    pre code {{
      background-color: transparent;
      border: 0;
      display: inline;
      line-height: inherit;
      padding: 0;
    }}
    blockquote {{
      border-left: 0.25em solid {border};
      color: {muted};
      padding: 0 1em;
      margin: 0 0 16px 0;
    }}
    table {{ border-spacing: 0; border-collapse: collapse; margin-bottom: 16px; }}
    table th, table td {{ border: 1px solid {border}; padding: 6px 13px; }}
    table th {{ background-color: {code_bg}; font-weight: 600; }}
    table tr:nth-child(2n) {{ background-color: {code_bg}; }}
    ul, ol {{ padding-left: 2em; margin-bottom: 16px; }}
    li {{ margin-bottom: 0.25em; }}
    a {{ color: {link}; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    img {{ max-width: 100%; height: auto; }}
    hr {{ border: none; border-top: 1px solid {border}; height: 1px; margin: 24px 0; }}
    {custom_css}
  </style>
</head>
<body>
  {html_body}
</body>
</html>"""
