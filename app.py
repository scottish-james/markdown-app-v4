"""
Office to Markdown - Main Application

This is the main entry point for the Office to Markdown converter application.
It sets up the Streamlit UI and coordinates the different components.
"""

import os
import streamlit as st
from src.converters.file_converter import convert_file_to_markdown
from src.converters.url_converter import convert_url_to_markdown
from src.processors.folder_processor import process_folder
from src.ui.components import setup_sidebar, get_supported_formats, setup_page_config
from config import UI_THEME_COLOR, APP_TITLE, APP_ICON

# Set up page configuration
setup_page_config()


def main():
    """Main application function that sets up the Streamlit interface."""

    # Initialize session state variables
    initialize_session_state()

    # App header
    st.title("Office to Markdown Converter")
    st.write("Convert your documents or websites to clean, structured Markdown with hyperlink extraction")

    # Set up the sidebar and get user preferences
    enhance_markdown, openai_api_key = setup_sidebar()

    # Set OpenAI API key in environment if provided
    if openai_api_key:
        os.environ["OPENAI_API_KEY"] = openai_api_key

    # Main content area - Tabs
    tab1, tab2, tab3 = st.tabs(["File Upload", "Website URL", "Folder Processing"])

    with tab1:
        handle_file_upload(enhance_markdown, openai_api_key)

    with tab2:
        handle_url_conversion(enhance_markdown, openai_api_key)

    with tab3:
        handle_folder_processing(enhance_markdown, openai_api_key)

    # Display output section if markdown content exists
    display_output_section()


def initialize_session_state():
    """Initialize Streamlit session state variables."""
    if "markdown_content" not in st.session_state:
        st.session_state.markdown_content = ""
    if "file_name" not in st.session_state:
        st.session_state.file_name = ""
    if "url_title" not in st.session_state:
        st.session_state.url_title = ""
    if "folder_processing_results" not in st.session_state:
        st.session_state.folder_processing_results = None


def handle_file_upload(enhance_markdown, openai_api_key):
    """Handle file upload and conversion logic."""
    # Get supported file extensions
    all_extensions = []
    formats = get_supported_formats()
    for category, info in formats.items():
        all_extensions.extend(info["extensions"])

    st.info(
        "Special feature: PDF and PowerPoint files will have their hyperlinks extracted and included in the markdown output."
    )

    uploaded_file = st.file_uploader(
        "Select a file to convert",
        type=all_extensions,
        help="Choose a file to convert to Markdown"
    )

    # Convert button for file
    if st.button("Convert File to Markdown", key="convert_file"):
        if not uploaded_file:
            st.error("Please upload a file to convert")
        else:
            with st.spinner("Converting to Markdown..."):
                # Show progress bar
                progress_bar = st.progress(0)
                for percent_complete in range(100):
                    progress_bar.progress(percent_complete + 1)

                # Convert uploaded file
                markdown_content, error = convert_file_to_markdown(
                    uploaded_file.getbuffer(),
                    uploaded_file.name,
                    enhance=enhance_markdown,
                    api_key=openai_api_key
                )

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = uploaded_file.name
                    st.success("Conversion completed successfully!")


def handle_url_conversion(enhance_markdown, openai_api_key):
    """Handle website URL conversion logic."""
    website_url = st.text_input(
        "Enter website URL",
        placeholder="https://example.com",
        help="Enter a complete URL to convert its content to Markdown"
    )

    # Convert button for URL
    if st.button("Convert Website to Markdown", key="convert_url"):
        if not website_url:
            st.error("Please enter a website URL to convert")
        else:
            with st.spinner("Fetching website and converting to Markdown..."):
                # Show progress bar
                progress_bar = st.progress(0)
                for percent_complete in range(100):
                    progress_bar.progress(percent_complete + 1)

                # Convert website URL
                markdown_content, error = convert_url_to_markdown(
                    website_url,
                    enhance=enhance_markdown,
                    api_key=openai_api_key
                )

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = f"{st.session_state.url_title}.md"
                    st.success("Website conversion completed successfully!")


def handle_folder_processing(enhance_markdown, openai_api_key):
    """Handle batch folder processing logic."""
    st.info(
        "Convert all supported files in a folder to markdown format. Each file will be processed and a .md file will be created."
    )

    # Input folder selection
    input_folder = st.text_input(
        "Enter path to folder with files to convert",
        placeholder="C:/Documents/MyFiles",
        help="Full path to the folder containing files to convert"
    )

    # Output folder selection (optional)
    output_folder = st.text_input(
        "Enter path to save markdown files (optional)",
        placeholder="Leave empty to create 'markdown' subfolder",
        help="Full path to save the converted markdown files. If left empty, files will be saved in a 'markdown' subfolder."
    )

    # Process folder button
    if st.button("Process Folder", key="process_folder"):
        if not input_folder or not os.path.isdir(input_folder):
            st.error("Please enter a valid folder path")
        else:
            # Process the folder
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # Create a generator to process files and update progress
                folder_processor = process_folder(
                    input_folder,
                    output_folder,
                    enhance=enhance_markdown,
                    api_key=openai_api_key
                )

                # Process files with progress updates
                for progress, status in folder_processor:
                    progress_bar.progress(min(1.0, progress))
                    status_text.text(status)

                # Get final results
                success_count, error_count, errors = next(folder_processor)

                # Save results to session state
                st.session_state.folder_processing_results = {
                    "success_count": success_count,
                    "error_count": error_count,
                    "errors": errors,
                    "output_folder": output_folder if output_folder else os.path.join(input_folder, "markdown")
                }

                # Show success message
                if success_count > 0:
                    st.success(f"Successfully converted {success_count} files.")

                # Show error message if any
                if error_count > 0:
                    st.warning(f"Failed to convert {error_count} files. See details below.")

            except Exception as e:
                st.error(f"Error processing folder: {str(e)}")

    # Display folder processing results if available
    display_folder_results()


def display_folder_results():
    """Display folder processing results if available."""
    if st.session_state.folder_processing_results:
        results = st.session_state.folder_processing_results

        st.subheader("Folder Processing Results")
        st.markdown(f"**Output Location:** {results['output_folder']}")
        st.markdown(f"**Successfully Converted:** {results['success_count']} files")
        st.markdown(f"**Failed Conversions:** {results['error_count']} files")

        # Show errors if any
        if results['error_count'] > 0:
            with st.expander("View Conversion Errors"):
                for file_name, error in results['errors'].items():
                    st.markdown(f"**{file_name}**: {error}")

        # Option to open output folder
        if st.button("Open Output Folder in File Explorer"):
            try:
                # Try to open the folder based on platform
                output_dir = results['output_folder']
                if os.path.exists(output_dir):
                    if os.name == 'nt':  # Windows
                        os.startfile(output_dir)
                    elif os.name == 'posix':  # macOS, Linux
                        import subprocess
                        if os.path.exists('/usr/bin/open'):  # macOS
                            subprocess.call(['open', output_dir])
                        else:  # Linux
                            subprocess.call(['xdg-open', output_dir])
                    st.success(f"Opened folder: {output_dir}")
                else:
                    st.error("Output folder not found")
            except Exception as e:
                st.error(f"Failed to open folder: {str(e)}")


def display_output_section():
    """Display the markdown output section if content exists."""
    if st.session_state.markdown_content:
        st.subheader("Converted Markdown")
        st.text_area(
            "Markdown Content",
            value=st.session_state.markdown_content,
            height=400
        )

        # Download button
        st.download_button(
            label="Download Markdown",
            data=st.session_state.markdown_content,
            file_name=st.session_state.file_name.rsplit(".", 1)[0] + ".md"
            if "." in st.session_state.file_name else st.session_state.file_name + ".md",
            mime="text/markdown"
        )


if __name__ == "__main__":
    main()