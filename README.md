# Markdown Editor with Live Preview

A modern, cross-platform desktop markdown editor with real-time HTML preview, built using PySide6 (Qt6 for Python).

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![PySide6](https://img.shields.io/badge/PySide6-6.8%2B-green)
![License](https://img.shields.io/badge/license-MIT-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Features

### Core Features

- **Split-window interface**: Text editor on the left, live HTML preview on the right
- **Real-time preview**: Updates automatically as you type (with a 300ms delay to prevent lag)
- **File operations**: New, Open, Save, Save As with standard keyboard shortcuts
- **Multi-format export**: Export to Markdown (.md), Plain Text (.txt), HTML (.html), Word (.docx), PDF (.pdf), RTF (.rtf), and ODT (.odt)
- **Keyboard shortcuts**: Standard shortcuts for file operations and editing (Ctrl/Cmd friendly)
- **Markdown extensions**: Support for code highlighting, tables, and table of contents
- **Clean interface**: Minimal, distraction-free design focused on writing and previewing
- **Cross-platform**: Works on Windows, macOS, and Linux
- **GitHub-style rendering**: Preview styled similar to GitHub's markdown rendering

### Smart Markdown Assistant (Unique Feature!)

This editor includes an intelligent Smart Assistant that provides real-time document analysis and quality insights.

#### Real-Time Document Analysis

- **Live Statistics**: Track word count, character count, line count, and estimated reading time
- **Structure Analysis**: Monitor headings (H1-H6), links, images, code blocks, lists, blockquotes, and tables
- **Quality Metrics**: Get instant readability scores and structure quality ratings
- **Issue Detection**: Automatically detect empty links, duplicate headings, and formatting problems
- **Heading Hierarchy**: Visual breakdown of your document's heading structure

#### Auto-Format Feature

- **One-Click Formatting**: Automatically format your entire document to markdown best practices
- **Smart Spacing**: Adds proper spacing before headings and after list markers
- **Consistency**: Ensures consistent formatting throughout your document
- **Clean Output**: Removes excessive blank lines and trailing whitespace

The Smart Assistant panel updates in real-time as you type (with an 800ms delay), providing continuous feedback on your document's quality and structure.

## Screenshots

_Screenshot coming soon - run the application to see the interface!_

## Installation

### Quick Start

```bash
# Clone the repository
git clone https://github.com/MichailSemoglou/smart-markdown-editor.git
cd smart-markdown-editor

# Create and activate virtual environment
python3 -m venv venv

# On macOS/Linux:
source venv/bin/activate

# On Windows:
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python markdown_editor.py
```

### Requirements

- Python 3.8 or higher
- PySide6 (>=6.8.0)
- Python Markdown (>=3.5.1)
- Optional export libraries (see below)

## Usage

Run the application:

```bash
python markdown_editor.py
```

### Keyboard Shortcuts

- `Ctrl+N` (or `Cmd+N` on Mac): New file
- `Ctrl+O` (or `Cmd+O` on Mac): Open file
- `Ctrl+S` (or `Cmd+S` on Mac): Save file
- `Ctrl+Shift+S` (or `Cmd+Shift+S` on Mac): Save As
- `Ctrl+Q` (or `Cmd+Q` on Mac): Exit
- `Ctrl+Z` (or `Cmd+Z` on Mac): Undo
- `Ctrl+Y` (or `Cmd+Y` on Mac): Redo

## Export Formats

The editor supports exporting your markdown documents to multiple formats:

- **Markdown (.md)**: Original markdown text
- **Plain Text (.txt)**: Plain text with markdown formatting removed
- **HTML (.html)**: Styled HTML with CSS
- **Word Document (.docx)**: Microsoft Word format (requires python-docx)
- **PDF Document (.pdf)**: Portable Document Format (requires reportlab or weasyprint)
- **Rich Text Format (.rtf)**: RTF format for compatibility
- **OpenDocument Text (.odt)**: ODT format for LibreOffice/OpenOffice

### Export Dependencies

Some export formats require additional libraries:

- **.docx**: `pip install python-docx`
- **.pdf**: `pip install reportlab` or `pip install weasyprint`

The application will gracefully handle missing dependencies and show helpful installation messages.

## Supported Markdown Features

- Headers (H1-H6)
- Emphasis (italic, bold)
- Lists (ordered and unordered)
- Links and images
- Code blocks and inline code
- Blockquotes
- Tables
- Horizontal rules
- And more standard markdown features

## Technical Details

- **GUI Framework**: PySide6 (Qt6 bindings for Python)
- **Markdown Processing**: Python Markdown library with extensions (codehilite, tables, toc)
- **HTML Preview**: QtWebEngineWidgets for rendering
- **Live Update**: QTimer-based delayed updates (300ms) for smooth typing experience
- **Architecture**: Model-View pattern with signal/slot connections

## Project Structure

```
smart-markdown-editor/
├── markdown_editor.py    # Main application file
├── requirements.txt      # Python dependencies
├── README.md            # Project documentation
├── LICENSE              # MIT License
├── .gitignore           # Git ignore rules
├── test_exports.py      # Test script for export functionality
└── test_sample.md       # Sample markdown file for testing
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Known Issues

- PDF export quality may vary depending on the PDF library used (reportlab vs weasyprint)
- Some complex markdown tables may not render perfectly in all export formats

## Future Enhancements

- [ ] Syntax highlighting in the editor
- [ ] Custom themes (dark mode)
- [ ] Find and replace functionality
- [ ] Custom CSS for preview
- [ ] Auto-save functionality
- [ ] Recent files menu

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [PySide6](https://wiki.qt.io/Qt_for_Python)
- Markdown processing by [Python Markdown](https://python-markdown.github.io/)
- Inspired by various markdown editors in the open-source community

## Support

If you encounter any issues or have questions, please file an issue on the [GitHub issue tracker](https://github.com/MichailSemoglou/smart-markdown-editor/issues).

---

Created as an educational project for learning GUI development with Python and Qt.
