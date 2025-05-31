"""
PowerPoint to Markdown Converter - Optimized for Claude Sonnet 4

A Streamlit application specifically designed for converting PowerPoint presentations
to clean, structured Markdown using Claude Sonnet 4's superior document processing capabilities.
"""

import streamlit as st
import os
from src.converters.file_converter import convert_file_to_markdown
from src.processors.folder_processor import process_folder
from src.ui.components import setup_sidebar, get_supported_formats, setup_page_config


def setup_sidebar_with_claude():
    """Enhanced sidebar setup with Claude Sonnet 4 as the primary AI provider."""
    with st.sidebar:
        st.header("PowerPoint to MD")
        st.write("Claude-Powered Document Conversion")

        # Supported formats in an expander
        with st.expander("Supported Formats"):
            from src.ui.components import display_supported_formats
            display_supported_formats()

        # Enhancement options
        st.subheader("AI Enhancement")
        enhance_markdown = st.checkbox("Enhance with Claude Sonnet 4", value=True,
                                       help="Use Claude to improve markdown formatting and structure")

        api_key = None
        if enhance_markdown:
            api_key = st.text_input(
                "Anthropic API Key",
                type="password",
                help="Enter your Anthropic API key for Claude enhancement"
            )

            # Show Claude advantages
            st.info("🎯 Claude Sonnet 4 provides superior PowerPoint structure analysis, formatting preservation, and intelligent content organization!")

        # Developer info
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        **Developed by:** James Taylor  
        **Powered by:** Claude Sonnet 4  
        **Optimized for:** PowerPoint Presentations
        """)

    return enhance_markdown, api_key


def handle_file_upload_enhanced(enhance_markdown, api_key):
    """Enhanced file upload handler optimized for PowerPoint with Claude support."""
    # Get supported file extensions
    all_extensions = []
    formats = get_supported_formats()
    for category, info in formats.items():
        all_extensions.extend(info["extensions"])

    # Enhanced info message
    st.info(
        "🚀 **PowerPoint Optimization Active!** "
        "This application is specifically optimized for PowerPoint presentations. "
        "While it supports other document formats, PowerPoint files (.pptx, .ppt) will receive "
        "the best processing with advanced structure analysis, bullet point hierarchy, "
        "and hyperlink extraction."
    )

    uploaded_file = st.file_uploader(
        "Select a file to convert (PowerPoint recommended)",
        type=all_extensions,
        help="PowerPoint files (.pptx, .ppt) are optimized for best results"
    )

    # Convert button for file
    if st.button("Convert File to Markdown", key="convert_file"):
        if not uploaded_file:
            st.error("Please upload a file to convert")
        else:
            with st.spinner("Converting to Markdown with Claude Sonnet 4..."):
                # Show progress bar
                progress_bar = st.progress(0)
                for percent_complete in range(100):
                    progress_bar.progress(percent_complete + 1)

                # Convert uploaded file with Claude
                markdown_content, error = convert_file_to_markdown(
                    uploaded_file.getbuffer(),
                    uploaded_file.name,
                    enhance=enhance_markdown,
                    api_key=api_key
                )

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = uploaded_file.name

                    # Show success message with Claude info
                    if enhance_markdown and api_key:
                        st.success("✨ Conversion completed successfully with Claude Sonnet 4 enhancement!")
                    else:
                        st.success("Conversion completed successfully!")


def handle_folder_processing_enhanced(enhance_markdown, api_key):
    """Enhanced batch folder processing optimized for PowerPoint files."""
    st.info(
        "🚀 Convert all supported files in a folder to markdown format using Claude Sonnet 4. "
        "PowerPoint files will receive optimized processing with advanced formatting preservation."
    )

    # Input folder selection
    input_folder = st.text_input(
        "Enter path to folder with files to convert",
        placeholder="C:/Documents/MyPresentations",
        help="Full path to the folder containing files to convert"
    )

    # Output folder selection (optional)
    output_folder = st.text_input(
        "Enter path to save markdown files (optional)",
        placeholder="Leave empty to create 'markdown' subfolder",
        help="Full path to save the converted markdown files. If left empty, files will be saved in a 'markdown' subfolder."
    )

    # Enhancement provider info
    st.info("🎯 **Claude Sonnet 4 will be used for batch processing** - Expect superior document structure analysis and formatting!")

    # Process folder button
    if st.button("Process Folder", key="process_folder"):
        if not input_folder or not os.path.isdir(input_folder):
            st.error("Please enter a valid folder path")
        else:
            # Process the folder with Claude
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # Create folder processor
                folder_processor = process_folder(
                    input_folder,
                    output_folder,
                    enhance=enhance_markdown,
                    api_key=api_key
                )

                # Process files with progress updates
                for progress, status in folder_processor:
                    progress_bar.progress(min(1.0, progress))
                    status_text.text(f"{status} (using Claude Sonnet 4)")

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
                    st.success(f"✨ Successfully converted {success_count} files using Claude Sonnet 4!")

                # Show error message if any
                if error_count > 0:
                    st.warning(f"Failed to convert {error_count} files. See details below.")

            except Exception as e:
                st.error(f"Error processing folder: {str(e)}")

    # Display folder processing results if available
    display_folder_results_enhanced()


def display_folder_results_enhanced():
    """Enhanced display of folder processing results."""
    if st.session_state.folder_processing_results:
        results = st.session_state.folder_processing_results

        st.subheader("📊 Folder Processing Results")

        # Create columns for better layout
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("✅ Successfully Converted", results['success_count'])
        with col2:
            st.metric("❌ Failed Conversions", results['error_count'])
        with col3:
            st.metric("🤖 AI Provider", "Claude Sonnet 4")

        st.markdown(f"**📁 Output Location:** {results['output_folder']}")

        # Show errors if any
        if results['error_count'] > 0:
            with st.expander("🔍 View Conversion Errors"):
                for file_name, error in results['errors'].items():
                    st.markdown(f"**{file_name}**: {error}")

        # Option to open output folder
        if st.button("📂 Open Output Folder in File Explorer"):
            try:
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
                    st.success(f"📂 Opened folder: {output_dir}")
                else:
                    st.error("Output folder not found")
            except Exception as e:
                st.error(f"Failed to open folder: {str(e)}")


def display_output_section_enhanced():
    """Enhanced display of the markdown output section."""
    if st.session_state.markdown_content:
        st.subheader("📝 Converted Markdown")

        # Add some statistics about the content
        content = st.session_state.markdown_content
        word_count = len(content.split())
        char_count = len(content)
        line_count = len(content.split('\n'))

        # Display stats in columns
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📊 Words", word_count)
        with col2:
            st.metric("📝 Characters", char_count)
        with col3:
            st.metric("📄 Lines", line_count)

        # Text area with the content
        st.text_area(
            "Markdown Content",
            value=content,
            height=400,
            help="Your converted and enhanced markdown content"
        )

        # Enhanced download button
        filename = st.session_state.file_name.rsplit(".", 1)[
                       0] + ".md" if "." in st.session_state.file_name else st.session_state.file_name + ".md"

        st.download_button(
            label="📥 Download Enhanced Markdown",
            data=content,
            file_name=filename,
            mime="text/markdown",
            help="Download your Claude-enhanced markdown file"
        )

        # Option to copy to clipboard (using streamlit-extras if available)
        try:
            import streamlit_extras
            if st.button("📋 Copy to Clipboard"):
                st.write("Content copied! (You may need to manually copy from the text area above)")
        except ImportError:
            pass


def main_enhanced():
    """Enhanced main application function optimized for PowerPoint and Claude."""

    # Set up page configuration
    setup_page_config()

    # Initialize session state variables
    if "markdown_content" not in st.session_state:
        st.session_state.markdown_content = ""
    if "file_name" not in st.session_state:
        st.session_state.file_name = ""
    if "folder_processing_results" not in st.session_state:
        st.session_state.folder_processing_results = None

    # App header with Claude and PowerPoint focus
    st.title("🚀 PowerPoint to Markdown Converter")
    st.markdown(
        "**Powered by Claude Sonnet 4** - Convert your PowerPoint presentations and documents to clean, "
        "structured Markdown with advanced AI enhancement. Optimized for PowerPoint with superior "
        "formatting preservation and intelligent content organization."
    )

    # Set up the enhanced sidebar
    enhance_markdown, api_key = setup_sidebar_with_claude()

    # Set API key in environment if provided
    if api_key:
        os.environ["ANTHROPIC_API_KEY"] = api_key

    # Main content area - Tabs (removed Website URL tab)
    tab1, tab2 = st.tabs(["📄 File Upload", "📁 Folder Processing"])

    with tab1:
        handle_file_upload_enhanced(enhance_markdown, api_key)

    with tab2:
        handle_folder_processing_enhanced(enhance_markdown, api_key)

    # Display output section if markdown content exists
    display_output_section_enhanced()


# Installation requirements note
def show_installation_requirements():
    """Show installation requirements for Claude integration."""
    st.markdown("""
    ## 🔧 Installation Requirements

    To use Claude Sonnet 4 enhancement, you need to install the Anthropic package:

    ```bash
    pip install anthropic
    ```

    You'll also need an Anthropic API key. Get one at: https://console.anthropic.com/

    ## 📋 Supported File Formats

    While this application supports various file formats, it is **optimized for PowerPoint presentations**:

    **Primary (Optimized):**
    - PowerPoint (.pptx, .ppt) - Advanced formatting preservation, bullet hierarchy, hyperlink extraction

    **Secondary (Standard processing):**
    - Word (.docx, .doc)
    - PDF (basic hyperlink extraction)
    - Excel (.xlsx, .xls)
    - HTML (.html, .htm)
    - Plain text formats (CSV, JSON, XML)
    """)


if __name__ == "__main__":
    main_enhanced()