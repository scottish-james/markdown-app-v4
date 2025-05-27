"""
Updated file converter module using MarkItDown's stream-based conversion.
"""

import io
from markitdown import MarkItDown
from openai import OpenAI
import os
from src.utils.file_utils import get_file_extension
from src.converters.hyperlink_extractor import extract_pdf_hyperlinks, extract_pptx_hyperlinks, \
    format_hyperlinks_section
from config import OPENAI_MODEL, OPENAI_TEMPERATURE, OPENAI_MAX_TOKENS, MARKDOWN_ENHANCEMENT_PROMPT


def enhance_markdown_formatting(markdown_content, api_key=None):
    """Enhance markdown formatting using OpenAI's model."""
    # [This function remains unchanged]
    try:
        if not api_key:
            # Try to get API key from environment or Streamlit secrets
            api_key = os.getenv("OPENAI_API_KEY")

        if not api_key:
            return markdown_content, "No OpenAI API key provided"

        # Initialize OpenAI client
        client = OpenAI(api_key=api_key)

        # Create a prompt for markdown enhancement
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


def convert_file_to_markdown(file_data, filename, enhance=True, api_key=None):
    """
    Convert a file to Markdown using MarkItDown and optionally enhance with AI.
    Uses enhanced PowerPoint processing for PPTX files.

    Args:
        file_data (bytes): The binary content of the file
        filename (str): The name of the file
        enhance (bool): Whether to enhance the markdown with AI
        api_key (str): OpenAI API key (optional)

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        ext = get_file_extension(filename)
        
        # Use enhanced processing for PowerPoint files
        if ext.lower() in ["pptx", "ppt"]:
            return convert_pptx_enhanced(file_data, filename, enhance, api_key)
        
        # Use standard MarkItDown for other file types
        return convert_standard_markitdown(file_data, filename, enhance, api_key)

    except Exception as e:
        return "", str(e)


def convert_pptx_enhanced(file_data, filename, enhance=True, api_key=None):
    """
    Convert PowerPoint files using enhanced processing that preserves formatting.
    """
    try:
        import tempfile
        from src.converters.enhanced_pptx_processor import convert_pptx_to_markdown_enhanced
        
        # Create temporary file for processing
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{get_file_extension(filename)}") as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name

        try:
            # Use enhanced PowerPoint processor
            markdown_content = convert_pptx_to_markdown_enhanced(tmp_file_path)
            
            # Extract and append hyperlinks
            hyperlinks = extract_pptx_hyperlinks(tmp_file_path)
            if hyperlinks:
                markdown_content += format_hyperlinks_section(hyperlinks, "Presentation")
            
        finally:
            # Clean up temporary file
            os.unlink(tmp_file_path)

        # Enhance with AI if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_formatting(markdown_content, api_key)
            if enhance_error:
                print(f"Enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None
        
    except Exception as e:
        return "", str(e)


def convert_standard_markitdown(file_data, filename, enhance=True, api_key=None):
    """
    Convert files using standard MarkItDown processing.
    """
    try:
        # Create a BytesIO object from the file data
        file_stream = io.BytesIO(file_data)
        file_stream.name = filename  # Set name attribute for MIME type detection

        # Initialize MarkItDown - not specifying enable_plugins to use default
        md = MarkItDown()

        # Try using direct file path conversion first (like in the working example)
        try:
            # Create a temporary file to use with the convert method
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{get_file_extension(filename)}") as tmp_file:
                tmp_file.write(file_data)
                tmp_file_path = tmp_file.name

            # Use the convert method with the file path (matching working example)
            result = md.convert(tmp_file_path)
            os.unlink(tmp_file_path)  # Delete temp file
        except Exception as file_path_error:
            # If file path method fails, fall back to convert_stream
            print(f"File path conversion failed: {str(file_path_error)}. Trying stream conversion...")
            # Reset the stream position before using convert_stream
            file_stream.seek(0)
            result = md.convert_stream(file_stream)

        # Try both result.markdown and result.text_content
        try:
            markdown_content = result.markdown
        except AttributeError:
            try:
                markdown_content = result.text_content
                print("Using result.text_content instead of result.markdown")
            except AttributeError:
                # If neither attribute exists, check what attributes are available
                print(f"Available attributes on result: {dir(result)}")
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
                    # Clean up
                    os.unlink(tmp_file_path)

        # Enhance with AI if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_formatting(markdown_content, api_key)
            if enhance_error:
                print(f"Enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        return markdown_content, None

    except Exception as e:
        return "", str(e)


def convert_stream_to_markdown(file_stream, filename, enhance=True, api_key=None):
    """
    Convert a file stream directly to Markdown.
    Use this when you already have a file-like object.

    Args:
        file_stream (io.BytesIO): A file-like object in binary mode
        filename (str): The name of the file (for extension detection)
        enhance (bool): Whether to enhance the markdown with AI
        api_key (str): OpenAI API key (optional)

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        # Save the current position
        current_pos = file_stream.tell()

        # Rewind to the beginning
        file_stream.seek(0)

        # Read the entire file data to use with convert_file_to_markdown
        file_data = file_stream.read()

        # Restore the stream position
        file_stream.seek(current_pos)

        # Use the other function to avoid code duplication
        return convert_file_to_markdown(file_data, filename, enhance, api_key)

    except Exception as e:
        return "", str(e)