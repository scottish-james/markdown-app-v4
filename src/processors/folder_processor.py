"""
Folder Processor Module

This module handles batch processing of files in a folder.
"""

import os
import glob
from src.converters.file_converter import convert_file_to_markdown
from src.utils.file_utils import get_supported_formats, get_all_supported_extensions, ensure_directory_exists
from config import DEFAULT_MARKDOWN_SUBFOLDER


def process_folder(folder_path, output_folder=None, enhance=True, api_key=None):
    """
    Process all compatible files in a folder and convert them to markdown.

    Args:
        folder_path (str): Path to folder containing files to convert
        output_folder (str): Path to save markdown files (defaults to subfolder in input folder)
        enhance (bool): Whether to enhance markdown with AI
        api_key (str): OpenAI API key

    Yields:
        Various progress updates and final results
    """
    # Setup output folder
    if not output_folder:
        output_folder = os.path.join(folder_path, DEFAULT_MARKDOWN_SUBFOLDER)

    # Create output folder if it doesn't exist
    ensure_directory_exists(output_folder)

    # Get all supported file extensions
    extensions = get_all_supported_extensions()

    # Track results
    success_count = 0
    error_count = 0
    errors = {}

    # Get all compatible files
    files_to_process = []
    for ext in extensions:
        files_to_process.extend(glob.glob(os.path.join(folder_path, f"*.{ext}")))

    total_files = len(files_to_process)

    # Early exit if no files found
    if total_files == 0:
        yield 1.0, "No compatible files found in folder"
        yield 0, 0, {}
        return

    # Process each file
    for i, file_path in enumerate(files_to_process):
        file_name = os.path.basename(file_path)

        try:
            # Update progress
            progress = (i + 1) / total_files
            yield progress, f"Processing {file_name} ({i + 1}/{total_files})"

            # Read file content
            with open(file_path, 'rb') as file:
                file_data = file.read()

            # Convert to markdown
            markdown_content, error = convert_file_to_markdown(
                file_data,
                file_name,
                enhance=enhance,
                api_key=api_key
            )

            if error:
                error_count += 1
                errors[file_name] = error
                continue

            # Save markdown content
            output_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.md")
            with open(output_file, 'w', encoding='utf-8') as md_file:
                md_file.write(markdown_content)

            success_count += 1

        except Exception as e:
            error_count += 1
            errors[file_name] = str(e)

    # Return final results
    yield success_count, error_count, errors


def find_compatible_files(directory_path):
    """
    Find all compatible files in a directory.

    Args:
        directory_path (str): The directory to search

    Returns:
        dict: Dictionary of files by category
    """
    formats = get_supported_formats()
    result = {}

    # Initialize result categories
    for category in formats.keys():
        result[category] = []

    # Only process if directory exists
    if not os.path.isdir(directory_path):
        return result

    # Find all files by extension
    for category, info in formats.items():
        for ext in info["extensions"]:
            files = glob.glob(os.path.join(directory_path, f"*.{ext}"))
            for file_path in files:
                result[category].append({
                    "path": file_path,
                    "name": os.path.basename(file_path),
                    "extension": ext
                })

    return result