"""
Updated file converter module that supports both OpenAI and Claude Sonnet 4 enhancement.
"""

import io
import os
from markitdown import MarkItDown
from openai import OpenAI
import anthropic
from typing import Tuple, Optional

from src.utils.file_utils import get_file_extension
from src.converters.hyperlink_extractor import extract_pdf_hyperlinks, extract_pptx_hyperlinks, \
    format_hyperlinks_section
from config import OPENAI_MODEL, OPENAI_TEMPERATURE, OPENAI_MAX_TOKENS, MARKDOWN_ENHANCEMENT_PROMPT

# Import the Claude enhancer (assuming it's in the same directory)
try:
    from .claude_markdown_enhancer import ClaudeMarkdownEnhancer, DOCUMENT_TO_MARKDOWN_SYSTEM_PROMPT

    CLAUDE_AVAILABLE = True
except ImportError:
    CLAUDE_AVAILABLE = False
    print("Claude enhancer not available. Install anthropic package and ensure claude_markdown_enhancer.py is present.")


def enhance_markdown_with_openai(markdown_content: str, api_key: str) -> Tuple[str, Optional[str]]:
    """Enhance markdown formatting using OpenAI's model."""
    try:
        if not api_key:
            return markdown_content, "No OpenAI API key provided"

        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": MARKDOWN_ENHANCEMENT_PROMPT},
                {"role": "user", "content": f"Please enhance this markdown content:\n\n{markdown_content}"}
            ],
            temperature=OPENAI_TEMPERATURE,
            max_tokens=OPENAI_MAX_TOKENS
        )

        enhanced_content = response.choices[0].message.content
        return enhanced_content, None

    except Exception as e:
        return markdown_content, str(e)


def enhance_markdown_with_claude(markdown_content: str, api_key: str,
                                 source_filename: str = "unknown",
                                 content_type: str = "Document") -> Tuple[str, Optional[str]]:
    """Enhance markdown formatting using Claude Sonnet 4."""
    if not CLAUDE_AVAILABLE:
        return markdown_content, "Claude enhancer not available"

    try:
        enhancer = ClaudeMarkdownEnhancer(api_key)
        return enhancer.enhance_markdown(markdown_content, source_filename, content_type)
    except Exception as e:
        return markdown_content, str(e)


def enhance_markdown_formatting(markdown_content: str, api_key: str = None,
                                provider: str = "claude",
                                source_filename: str = "unknown",
                                content_type: str = "Document") -> Tuple[str, Optional[str]]:
    """
    Enhance markdown formatting using either OpenAI or Claude.

    Args:
        markdown_content (str): The markdown content to enhance
        api_key (str): API key for the chosen provider
        provider (str): Either "openai" or "claude"
        source_filename (str): Source filename for context (used by Claude)
        content_type (str): Document type (used by Claude)

    Returns:
        Tuple[str, Optional[str]]: Enhanced content and error message (if any)
    """
    if not api_key:
        # Try to get API key from environment
        if provider == "claude":
            api_key = os.getenv("ANTHROPIC_API_KEY")
        else:
            api_key = os.getenv("OPENAI_API_KEY")

    if not api_key:
        return markdown_content, f"No {provider.upper()} API key provided"

    if provider.lower() == "claude":
        return enhance_markdown_with_claude(markdown_content, api_key, source_filename, content_type)
    else:
        return enhance_markdown_with_openai(markdown_content, api_key)


def convert_file_to_markdown(file_data, filename, enhance=True, api_key=None,
                             enhancement_provider="claude"):
    """
    Convert a file to Markdown using MarkItDown and optionally enhance with AI.

    Args:
        file_data (bytes): The binary content of the file
        filename (str): The name of the file
        enhance (bool): Whether to enhance the markdown with AI
        api_key (str): API key for the enhancement provider
        enhancement_provider (str): Either "openai" or "claude"

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        ext = get_file_extension(filename)

        # Use enhanced processing for PowerPoint files
        if ext.lower() in ["pptx", "ppt"]:
            return convert_pptx_enhanced(file_data, filename, enhance, api_key, enhancement_provider)

        # Use standard MarkItDown for other file types
        return convert_standard_markitdown(file_data, filename, enhance, api_key, enhancement_provider)

    except Exception as e:
        return "", str(e)


def convert_pptx_enhanced(file_data, filename, enhance=True, api_key=None, enhancement_provider="claude"):
    """Convert PowerPoint files using enhanced processing that preserves formatting."""
    try:
        import tempfile
        from src.converters.enhanced_pptx_processor import convert_pptx_to_markdown_enhanced

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

        # Enhance with AI if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_formatting(
                markdown_content,
                api_key,
                enhancement_provider,
                filename,
                "PowerPoint Presentation"
            )
            if enhance_error:
                print(f"Enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None

    except Exception as e:
        return "", str(e)


def convert_standard_markitdown(file_data, filename, enhance=True, api_key=None, enhancement_provider="claude"):
    """Convert files using standard MarkItDown processing."""
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

        # Enhance with AI if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_formatting(
                markdown_content,
                api_key,
                enhancement_provider,
                filename,
                content_type
            )
            if enhance_error:
                print(f"Enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None

    except Exception as e:
        return "", str(e)


def convert_stream_to_markdown(file_stream, filename, enhance=True, api_key=None, enhancement_provider="claude"):
    """Convert a file stream directly to Markdown."""
    try:
        current_pos = file_stream.tell()
        file_stream.seek(0)
        file_data = file_stream.read()
        file_stream.seek(current_pos)

        return convert_file_to_markdown(file_data, filename, enhance, api_key, enhancement_provider)

    except Exception as e:
        return "", str(e)


# Example usage
if __name__ == "__main__":
    # Test with a sample markdown file
    sample_content = b"""# Test Document

This is a test document with some content.

* Item 1
* Item 2
  * Nested item

## Section 2
More content here."""

    # Test with Claude
    claude_api_key = os.getenv("ANTHROPIC_API_KEY")
    if claude_api_key:
        result, error = convert_file_to_markdown(
            sample_content,
            "test.md",
            enhance=True,
            api_key=claude_api_key,
            enhancement_provider="claude"
        )
        print("Claude Enhanced Result:")
        print("=" * 50)
        print(result)
        if error:
            print(f"Error: {error}")

    # Test with OpenAI
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if openai_api_key:
        result, error = convert_file_to_markdown(
            sample_content,
            "test.md",
            enhance=True,
            api_key=openai_api_key,
            enhancement_provider="openai"
        )
        print("\nOpenAI Enhanced Result:")
        print("=" * 50)
        print(result)
        if error:
            print(f"Error: {error}")