"""Pytest tests for export functionality.

These tests verify that each built-in exporter in :mod:`src.exporters.builtin`
produces a non-empty output file with the correct content.  Optional-library
exporters (docx, pdf) are skipped automatically when their dependency is absent.
"""

import importlib
from pathlib import Path

import markdown
import pytest

# ---------------------------------------------------------------------------
# Sample document used across all export tests
# ---------------------------------------------------------------------------

SAMPLE_MD = """\
# Test Document

This is a **bold** and _italic_ test.

## Features

- Bullet item one
- Bullet item two

1. Ordered item one
2. Ordered item two

### Code Example

```python
print("Hello, World!")
```

> This is a blockquote.

---

Plain paragraph at the end.
"""


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _is_available(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


# ---------------------------------------------------------------------------
# Library availability (non-export)
# ---------------------------------------------------------------------------

def test_markdown_conversion():
    """markdown library converts SAMPLE_MD without raising."""
    html = markdown.markdown(SAMPLE_MD, extensions=["codehilite", "tables", "toc"])
    assert "<h1 " in html
    assert "<h2 " in html
    assert "<ul>" in html


def test_html2text_available():
    pytest.importorskip("html2text")
    import html2text

    h = html2text.HTML2Text()
    result = h.handle("<p>Hello</p>")
    assert "Hello" in result


# ---------------------------------------------------------------------------
# Exporter integration tests
# ---------------------------------------------------------------------------

def test_markdown_exporter(tmp_path: Path):
    from src.exporters.builtin import MarkdownExporter

    out = tmp_path / "out.md"
    MarkdownExporter().export(SAMPLE_MD, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "# Test Document" in content


def test_text_exporter(tmp_path: Path):
    from src.exporters.builtin import TextExporter

    out = tmp_path / "out.txt"
    TextExporter().export(SAMPLE_MD, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    # Headings markers should be stripped
    assert "# Test Document" not in content
    assert "Test Document" in content


def test_html_exporter(tmp_path: Path):
    from src.exporters.builtin import HTMLExporter

    out = tmp_path / "out.html"
    HTMLExporter().export(SAMPLE_MD, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "<!DOCTYPE html>" in content
    assert "<h1 " in content


def test_rtf_exporter(tmp_path: Path):
    from src.exporters.builtin import RTFExporter

    out = tmp_path / "out.rtf"
    RTFExporter().export(SAMPLE_MD, out)
    assert out.exists()
    content = out.read_text(encoding="ascii", errors="replace")
    assert r"\rtf1" in content
    assert "Test Document" in content


def test_odt_exporter(tmp_path: Path):
    import zipfile

    from src.exporters.builtin import ODTExporter

    out = tmp_path / "out.odt"
    ODTExporter().export(SAMPLE_MD, out)
    assert out.exists()
    # ODT files must be valid ZIP archives
    assert zipfile.is_zipfile(out)
    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    assert "mimetype" in names
    assert "content.xml" in names
    assert "META-INF/manifest.xml" in names


@pytest.mark.skipif(
    not _is_available("docx"), reason="python-docx not installed"
)
def test_docx_exporter(tmp_path: Path):
    from docx import Document

    from src.exporters.builtin import DocxExporter

    out = tmp_path / "out.docx"
    DocxExporter().export(SAMPLE_MD, out)
    assert out.exists()
    doc = Document(str(out))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "Test Document" in full_text


@pytest.mark.skipif(
    not (_is_available("weasyprint") or _is_available("reportlab")),
    reason="neither weasyprint nor reportlab installed",
)
def test_pdf_exporter(tmp_path: Path):
    from src.exporters.builtin import PDFExporter

    out = tmp_path / "out.pdf"
    PDFExporter().export(SAMPLE_MD, out)
    assert out.exists()
    assert out.stat().st_size > 0
    # All PDF files start with the %PDF magic bytes
    assert out.read_bytes()[:4] == b"%PDF"


# ---------------------------------------------------------------------------
# Registry smoke-test
# ---------------------------------------------------------------------------

def test_register_all_populates_registry():
    from src.exporters import _EXPORTERS, get_available_exporters
    from src.exporters.builtin import register_all

    register_all()
    assert len(_EXPORTERS) >= 7
    # get_available_exporters must return at least the pure-Python ones
    available = get_available_exporters()
    extensions = {cls.extension for cls in available}
    assert {"md", "txt", "html", "rtf", "odt"}.issubset(extensions)

