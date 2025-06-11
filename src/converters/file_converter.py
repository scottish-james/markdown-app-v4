"""
File converter module optimized for Claude Sonnet 4 enhancement.
Specifically designed for PowerPoint presentations with advanced processing.
"""

import io
import os
from markitdown import MarkItDown
from typing import Tuple, Optional

from src.utils.file_utils import get_file_extension
from src.converters.hyperlink_extractor import extract_pdf_hyperlinks, format_hyperlinks_section

# Import the Claude enhancer
try:
    from .claude_markdown_convertor import ClaudeMarkdownEnhancer
    CLAUDE_AVAILABLE = True
except ImportError:
    CLAUDE_AVAILABLE = False
    print("Claude enhancer not available. Install anthropic package and ensure claude_markdown_convertor.py is present.")


def enhance_markdown_with_claude(markdown_content: str, api_key: str,
                                 source_filename: str = "unknown",
                                 content_type: str = "Document") -> Tuple[str, Optional[str]]:
    """
    Enhance markdown formatting using Claude Sonnet 4.

    Args:
        markdown_content (str): The markdown content to enhance
        api_key (str): Anthropic API key
        source_filename (str): Source filename for context
        content_type (str): Document type

    Returns:
        Tuple[str, Optional[str]]: Enhanced content and error message (if any)
    """
    if not CLAUDE_AVAILABLE:
        return markdown_content, "Claude enhancer not available"

    try:
        enhancer = ClaudeMarkdownEnhancer(api_key)
        return enhancer.enhance_markdown(markdown_content, source_filename, content_type)
    except Exception as e:
        return markdown_content, str(e)


def convert_file_to_markdown(file_data, filename, enhance=True, api_key=None):
    """
    Convert a file to Markdown using MarkItDown and enhance with Claude Sonnet 4.
    Optimized for PowerPoint presentations.

    Args:
        file_data (bytes): The binary content of the file
        filename (str): The name of the file
        enhance (bool): Whether to enhance the markdown with Claude
        api_key (str): Anthropic API key for Claude enhancement

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        ext = get_file_extension(filename)

        # Use enhanced processing for PowerPoint files (primary optimization)
        if ext.lower() in ["pptx", "ppt"]:
            return convert_pptx_enhanced(file_data, filename, enhance, api_key)

        # Use standard MarkItDown for other file types
        return convert_standard_markitdown(file_data, filename, enhance, api_key)

    except Exception as e:
        return "", str(e)


def convert_pptx_enhanced(file_data, filename, enhance=True, api_key=None):
    """
    Convert PowerPoint files using enhanced processing that preserves formatting.
    This is the core optimization of this application.
    """
    try:
        import tempfile
        from src.processors.enhanced_pptx_processor import convert_pptx_to_markdown_enhanced

        # Create temporary file for processing
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{get_file_extension(filename)}") as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name

        try:
            # Use enhanced PowerPoint processor (includes inline hyperlinks)
            markdown_content = convert_pptx_to_markdown_enhanced(tmp_file_path)
        finally:
            # Clean up temporary file
            os.unlink(tmp_file_path)

        # Enhance with Claude if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_with_claude(
                markdown_content,
                api_key,
                filename,
                "PowerPoint Presentation"
            )
            if enhance_error:
                print(f"Claude enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None

    except Exception as e:
        return "", str(e)


def convert_standard_markitdown(file_data, filename, enhance=True, api_key=None):
    """
    Convert files using standard MarkItDown processing with Claude enhancement.
    """
    try:
        # Create a BytesIO object from the file data
        file_stream = io.BytesIO(file_data)
        file_stream.name = filename

        # Initialize MarkItDown
        md = MarkItDown()

        # Convert using temporary file method (more reliable)
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{get_file_extension(filename)}") as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name

        try:
            result = md.convert(tmp_file_path)
            os.unlink(tmp_file_path)
        except Exception as file_path_error:
            print(f"File path conversion failed: {str(file_path_error)}. Trying stream conversion...")
            file_stream.seek(0)
            result = md.convert_stream(file_stream)

        # Get markdown content
        try:
            markdown_content = result.markdown
        except AttributeError:
            try:
                markdown_content = result.text_content
            except AttributeError:
                raise Exception("Neither 'markdown' nor 'text_content' attribute found on result object")

        # Extract hyperlinks for PDF files
        ext = get_file_extension(filename)
        if ext.lower() == "pdf":
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp_file:
                tmp_file.write(file_data)
                tmp_file_path = tmp_file.name

                try:
                    hyperlinks = extract_pdf_hyperlinks(tmp_file_path)
                    markdown_content += format_hyperlinks_section(hyperlinks, "Document")
                finally:
                    os.unlink(tmp_file_path)

        # Determine content type for Claude
        content_type_map = {
            "pdf": "PDF Document",
            "docx": "Word Document",
            "doc": "Word Document",
            "xlsx": "Excel Spreadsheet",
            "xls": "Excel Spreadsheet",
            "html": "HTML Document",
            "csv": "CSV File",
            "json": "JSON File",
            "xml": "XML File"
        }
        content_type = content_type_map.get(ext.lower(), "Document")

        # Enhance with Claude if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_with_claude(
                markdown_content,
                api_key,
                filename,
                content_type
            )
            if enhance_error:
                print(f"Claude enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None

    except Exception as e:
        return "", str(e)


def convert_stream_to_markdown(file_stream, filename, enhance=True, api_key=None):
    """Convert a file stream directly to Markdown with Claude enhancement."""
    try:
        current_pos = file_stream.tell()
        file_stream.seek(0)
        file_data = file_stream.read()
        file_stream.seek(current_pos)

        return convert_file_to_markdown(file_data, filename, enhance, api_key)

    except Exception as e:
        return "", str(e)