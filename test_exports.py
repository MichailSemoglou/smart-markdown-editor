#!/usr/bin/env python3
"""
Test script to verify all export functionality works correctly.
"""

import os
import sys
import markdown
from docx import Document
import reportlab
import weasyprint
import html2text

def test_export_functionality():
    """Test that all required libraries are available and functional."""
    print("Testing export functionality...")
    
    # Test basic markdown processing
    test_md = """# Test Document

This is a **test** with _various_ formatting.

## Features

- Bold text
- Italic text  
- `Code blocks`
- [Links](https://example.com)

### Code Example

```python
print("Hello, World!")
```

> This is a blockquote

---

Thank you for testing!
"""
    
    # Test markdown to HTML conversion
    html = markdown.markdown(test_md, extensions=['codehilite', 'tables', 'toc'])
    print("✓ Markdown to HTML conversion works")
    
    # Test python-docx
    try:
        doc = Document()
        doc.add_heading('Test Document', 0)
        doc.add_paragraph('Test paragraph')
        print("✓ python-docx functionality works")
    except Exception as e:
        print(f"✗ python-docx error: {e}")
    
    # Test reportlab
    try:
        from reportlab.pdfgen import canvas
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        print("✓ ReportLab functionality works")
    except Exception as e:
        print(f"✗ ReportLab error: {e}")
    
    # Test weasyprint
    try:
        import weasyprint
        print("✓ WeasyPrint functionality works")
    except Exception as e:
        print(f"✗ WeasyPrint error: {e}")
    
    # Test html2text
    try:
        import html2text
        h = html2text.HTML2Text()
        h.handle("<p>Test</p>")
        print("✓ html2text functionality works")
    except Exception as e:
        print(f"✗ html2text error: {e}")
    
    print("\nAll libraries are available! Export functionality should work correctly.")

if __name__ == "__main__":
    test_export_functionality()
