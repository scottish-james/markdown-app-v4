"""
File Converter Module

This module handles the conversion of various file types to Markdown format.
"""

import os
import tempfile
from markitdown import MarkItDown
from openai import OpenAI
from src.utils.file_utils import get_file_extension
from src.converters.hyperlink_extractor import extract_pptx_hyperlinks, extract_pdf_hyperlinks, format_hyperlinks_section
from config import OPENAI_MODEL, OPENAI_TEMPERATURE, OPENAI_MAX_TOKENS, MARKDOWN_ENHANCEMENT_PROMPT


def enhance_markdown_formatting(markdown_content, api_key=None):
    """
    Enhance markdown formatting using OpenAI's GPT-4o model.

    Args:
        markdown_content (str): The raw markdown content to enhance
        api_key (str): OpenAI API key (optional if stored in environment)

    Returns:
        tuple: (enhanced_markdown, error_message)
    """
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
    Special handling for PowerPoint and PDF files to extract hyperlinks.

    Args:
        file_data (bytes): The binary content of the file
        filename (str): The name of the file
        enhance (bool): Whether to enhance the markdown with AI
        api_key (str): OpenAI API key (optional)

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        # Create a temporary file
        ext = get_file_extension(filename)
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name

        # Initialize MarkItDown and convert the file
        md = MarkItDown(enable_plugins=False)
        result = md.convert(tmp_file_path)
        markdown_content = result.text_content

        # Special handling for PowerPoint files
        if ext.lower() in ["pptx", "ppt"]:
            # Extract hyperlinks from PowerPoint
            hyperlinks = extract_pptx_hyperlinks(tmp_file_path)
            # Add hyperlinks to the markdown content
            markdown_content += format_hyperlinks_section(hyperlinks, "Presentation")

        # Special handling for PDF files
        elif ext.lower() == "pdf":
            try:
                # Extract hyperlinks from PDF
                hyperlinks = extract_pdf_hyperlinks(tmp_file_path)
                # Add hyperlinks to the markdown content
                markdown_content += format_hyperlinks_section(hyperlinks, "Document")
            except Exception as e:
                print(f"PDF hyperlink extraction error: {str(e)}. Continuing with basic conversion.")

        # Clean up the temporary file
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