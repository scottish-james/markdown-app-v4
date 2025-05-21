"""
Tests for the URL converter module.
"""

import unittest
from unittest.mock import patch, MagicMock
import tempfile
import os
from src.converters.url_converter import convert_url_to_markdown


class TestURLConverter(unittest.TestCase):
    """Test cases for URL converter functionality."""

    @patch('src.converters.url_converter.requests.get')
    @patch('src.converters.url_converter.MarkItDown')
    @patch('src.converters.url_converter.BeautifulSoup')
    @patch('src.converters.url_converter.enhance_markdown_formatting')
    def test_convert_url_basic(self, mock_enhance, mock_bs, mock_markitdown, mock_get):
        """Test basic URL conversion without enhancement."""
        # Mock response
        mock_response = MagicMock()
        mock_response.content = b"<html><head><title>Test Page</title></head><body>Test content</body></html>"
        mock_get.return_value = mock_response

        # Mock BeautifulSoup
        mock_soup = MagicMock()
        mock_title = MagicMock()
        mock_title.string = "Test Page"
        mock_soup.title = mock_title
        mock_bs.return_value = mock_soup

        # Mock MarkItDown
        mock_md_instance = MagicMock()
        mock_markitdown.return_value = mock_md_instance
        mock_result = MagicMock()
        mock_result.text_content = "# Test Page\n\nTest content"
        mock_md_instance.convert.return_value = mock_result

        # Test URL conversion
        markdown, error, title = convert_url_to_markdown(
            "https://example.com",
            enhance=False
        )

        # Assertions
        self.assertIsNone(error)
        self.assertEqual(markdown, "# Test Page\n\nTest content")
        self.assertEqual(title, "Test_Page")
        mock_get.assert_called_once_with(
            "https://example.com",
            headers={"User-Agent": unittest.mock.ANY},
            timeout=10
        )
        mock_md_instance.convert.assert_called_once()

    @patch('src.converters.url_converter.requests.get')
    @patch('src.converters.url_converter.MarkItDown')
    @patch('src.converters.url_converter.BeautifulSoup')
    @patch('src.converters.url_converter.enhance_markdown_formatting')
    def test_convert_url_with_enhancement(self, mock_enhance, mock_bs, mock_markitdown, mock_get):
        """Test URL conversion with enhancement."""
        # Mock response
        mock_response = MagicMock()
        mock_response.content = b"<html><head><title>Test Page</title></head><body>Test content</body></html>"
        mock_get.return_value = mock_response

        # Mock BeautifulSoup
        mock_soup = MagicMock()
        mock_title = MagicMock()
        mock_title.string = "Test Page"
        mock_soup.title = mock_title
        mock_bs.return_value = mock_soup

        # Mock MarkItDown
        mock_md_instance = MagicMock()
        mock_markitdown.return_value = mock_md_instance
        mock_result = MagicMock()
        mock_result.text_content = "# Test Page\n\nTest content"
        mock_md_instance.convert.return_value = mock_result

        # Mock enhancement
        mock_enhance.return_value = ("# Enhanced Test Page\n\nTest content with formatting", None)

        # Test URL conversion with enhancement
        markdown, error, title = convert_url_to_markdown(
            "https://example.com",
            enhance=True,
            api_key="test_api_key"
        )

        # Assertions
        self.assertIsNone(error)
        self.assertEqual(markdown, "# Enhanced Test Page\n\nTest content with formatting")
        self.assertEqual(title, "Test_Page")
        mock_enhance.assert_called_once_with("# Test Page\n\nTest content", "test_api_key")

    @patch('src.converters.url_converter.requests.get')
    def test_invalid_url_format(self, mock_get):
        """Test conversion with invalid URL format."""
        # Test with invalid URL
        markdown, error, title = convert_url_to_markdown(
            "invalid-url"
        )

        # Assertions
        self.assertEqual(markdown, "")
        self.assertEqual(error, "Invalid URL format. Please include http:// or https://")
        self.assertEqual(title, "")
        mock_get.assert_not_called()

    @patch('src.converters.url_converter.requests.get')
    def test_request_exception(self, mock_get):
        """Test handling of request exceptions."""
        # Mock request exception
        mock_get.side_effect = Exception("Connection error")

        # Test with exception
        markdown, error, title = convert_url_to_markdown(
            "https://example.com"
        )

        # Assertions
        self.assertEqual(markdown, "")
        self.assertEqual(error, "Connection error")
        self.assertEqual(title, "")


if __name__ == '__main__':
    unittest.main()