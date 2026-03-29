"""Unit tests for the MarkdownAnalyzer.

This module tests the MarkdownAnalyzer class including:
- Word count (excluding code blocks)
- Heading detection
- Reading time estimation
- Empty document handling
- Issue detection (empty links, duplicate headings)
"""

import unittest
from src.core.analyzer import MarkdownAnalyzer, get_readability_color


class TestMarkdownAnalyzer(unittest.TestCase):
    """Test the markdown analyzer functionality."""
    
    def test_empty_document(self):
        """Test with empty document."""
        analyzer = MarkdownAnalyzer("")
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['word_count'], 0)
        self.assertEqual(metrics['char_count'], 0)
        self.assertEqual(metrics['line_count'], 1)  # Empty document has 1 line
        self.assertEqual(metrics['reading_time'], 1)
        
        # Check structure
        self.assertEqual(
            metrics['headings'], 
            {'h1': 0, 'h2': 0, 'h3': 0, 'h4': 0, 'h5': 0, 'h6': 0}
        )
        self.assertEqual(metrics['links'], 0)
        self.assertEqual(metrics['images'], 0)
        self.assertEqual(metrics['code_blocks'], 0)
        
        # Check quality
        self.assertEqual(metrics['readability_score'], 100)  # Empty gets max score
        self.assertEqual(metrics['structure_quality'], "No structure")
    
    def test_simple_document(self):
        """Test with simple markdown document."""
        analyzer = MarkdownAnalyzer("# Title\n\nParagraph text.")
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['word_count'], 3)  # Title, Paragraph, text
        self.assertEqual(metrics['char_count'], 24)
        self.assertEqual(metrics['line_count'], 3)
        
        # Check headings
        self.assertEqual(
            metrics['headings'], 
            {'h1': 1, 'h2': 0, 'h3': 0, 'h4': 0, 'h5': 0, 'h6': 0}
        )
    
    def test_word_count_excludes_code_blocks(self):
        """Test that code blocks are excluded from word count."""
        text = """# Title

```python
def example():
    pass
```

Paragraph text."""
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        # Code block should be excluded
        self.assertEqual(metrics['word_count'], 3)  # Title, Paragraph, text
        self.assertEqual(metrics['code_blocks'], 1)
    
    def test_heading_detection(self):
        """Test heading detection."""
        text = """# Heading 1

## Heading 2

### Heading 3
"""
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(
            metrics['headings'], 
            {'h1': 1, 'h2': 1, 'h3': 1, 'h4': 0, 'h5': 0, 'h6': 0}
        )
    
    def test_reading_time_estimation(self):
        """Test reading time estimation."""
        # 200 words = 1 minute
        text = " ".join(["word"] * 200)
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['reading_time'], 1)
        
        # 400 words = 2 minutes
        text = " ".join(["word"] * 400)
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['reading_time'], 2)
    
    def test_link_detection(self):
        """Test link detection."""
        text = "[Example Link](https://example.com)"
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['links'], 1)
    
    def test_image_detection(self):
        """Test image detection."""
        text = "![Alt text](image.png)"
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['images'], 1)
    
    def test_list_detection(self):
        """Test list detection."""
        text = """- Item 1
- Item 2
- Item 3
"""
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['lists'], 3)
    
    def test_blockquote_detection(self):
        """Test blockquote detection."""
        text = "> This is a quote"
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['blockquotes'], 1)
    
    def test_table_detection(self):
        """Test table detection."""
        text = """| Header 1 | Header 2 |
|---------|---------|
| Cell 1  | Cell 2  |"""
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        
        self.assertEqual(metrics['tables'], 1)
    
    def test_get_readability_color(self):
        """Test readability color helper function."""
        self.assertEqual(get_readability_color(90), "green")
        self.assertEqual(get_readability_color(70), "orange")
        self.assertEqual(get_readability_color(50), "red")
        self.assertEqual(get_readability_color(30), "red")


class TestInputValidation(unittest.TestCase):
    """Test input validation in analyzer."""

    def test_none_input(self):
        """Test that None input is handled gracefully (treated as empty string)."""
        analyzer = MarkdownAnalyzer(None)
        metrics = analyzer.analyze()
        self.assertEqual(metrics['word_count'], 0)
        self.assertEqual(metrics['char_count'], 0)
        self.assertEqual(metrics['structure_quality'], "No structure")

    def test_numeric_input(self):
        """Test that non-string input is converted to string."""
        analyzer = MarkdownAnalyzer(12345)
        metrics = analyzer.analyze()
        # "12345" is one word
        self.assertEqual(metrics['word_count'], 1)

    def test_unicode_content(self):
        """Test with unicode content."""
        analyzer = MarkdownAnalyzer("# Héllo\n\nWörld content.")
        metrics = analyzer.analyze()
        self.assertGreater(metrics['word_count'], 0)
        self.assertEqual(metrics['headings']['h1'], 1)

    def test_only_code_blocks(self):
        """Test document with only code blocks has zero word count."""
        text = "```python\nprint('hello')\n```"
        analyzer = MarkdownAnalyzer(text)
        metrics = analyzer.analyze()
        self.assertEqual(metrics['word_count'], 0)
        self.assertEqual(metrics['code_blocks'], 1)

    def test_headings_cache_consistency(self):
        """Test that _analyze_headings is consistent across multiple calls."""
        text = "# Title\n## Section\n### Sub"
        analyzer = MarkdownAnalyzer(text)
        first = analyzer._analyze_headings()
        second = analyzer._analyze_headings()
        self.assertIs(first, second)  # Same object — cached


if __name__ == "__main__":
    unittest.main()
