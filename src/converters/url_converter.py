"""
URL Converter Module

This module handles the conversion of web URLs to Markdown format.
"""

import os
import tempfile
import requests
from markitdown import MarkItDown
from bs4 import BeautifulSoup
from urllib.parse import urlparse
from src.converters.file_converter import enhance_markdown_formatting
from config import USER_AGENT


def convert_url_to_markdown(url, enhance=True, api_key=None):
    """
    Convert a website URL to Markdown by fetching HTML and using MarkItDown.

    Args:
        url (str): The website URL to convert
        enhance (bool): Whether to enhance the markdown with AI
        api_key (str): OpenAI API key (optional)

    Returns:
        tuple: (markdown_content, error_message)
    """
    try:
        # Validate URL format
        parsed_url = urlparse(url)
        if not parsed_url.scheme or not parsed_url.netloc:
            return "", "Invalid URL format. Please include http:// or https://"

        # Add scheme if missing
        if not parsed_url.scheme:
            url = f"https://{url}"

        # Fetch the HTML content
        headers = {
            "User-Agent": USER_AGENT
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()  # Raise exception for 4XX/5XX status codes

        # Save HTML to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
            tmp_file.write(response.content)
            tmp_file_path = tmp_file.name

        # Initialize MarkItDown and convert the HTML file
        md = MarkItDown(enable_plugins=False)
        result = md.convert(tmp_file_path)

        # Clean up the temporary file
        os.unlink(tmp_file_path)

        markdown_content = result.text_content

        # Get page title for filename suggestion
        try:
            soup = BeautifulSoup(response.content, 'html.parser')
            page_title = soup.title.string if soup.title else parsed_url.netloc
            # Clean the title for use as a filename
            page_title = "".join(c if c.isalnum() or c in " -_" else "_" for c in page_title).strip()
            if len(page_title) > 50:
                page_title = page_title[:50]
            # Return the title through the session state or another mechanism
            return_title = page_title
        except:
            return_title = parsed_url.netloc

        # Enhance with AI if enabled
        if enhance and api_key:
            enhanced_content, enhance_error = enhance_markdown_formatting(markdown_content, api_key)
            if enhance_error:
                print(f"Enhancement error: {enhance_error}. Using original markdown.")
            else:
                markdown_content = enhanced_content

        # Updated to also return the page title
        return markdown_content, None, return_title
    except requests.exceptions.RequestException as e:
        return "", f"Error fetching URL: {str(e)}", ""
    except Exception as e:
        return "", str(e), ""