"""
Document to Markdown Converter - Optimized for Claude Sonnet 4

A Streamlit application for converting various document formats
to clean, structured Markdown using Claude Sonnet 4's superior document processing capabilities.
"""

import streamlit as st
import os
from src.converters.file_converter import convert_file_to_markdown
from src.processors.folder_processor import process_folder
from src.ui.components import setup_sidebar, get_supported_formats, setup_page_config
from src.ui.about_tab import render_about_tab
from src.content.features import get_main_features, get_feature_tagline, get_tool_description


def setup_enhanced_sidebar():
    """Enhanced sidebar setup with Claude Sonnet 4 as the primary AI provider."""
    with st.sidebar:
        st.header("Document to Markdown")
        st.write("AI-Powered Document Conversion")

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

            # Clear instructions for getting API key
            if not api_key:
                st.info("💡 **Need an API key?**")
                st.markdown("""
                1. Visit [console.anthropic.com](https://console.anthropic.com/)
                2. Create an account or sign in
                3. Generate a new API key
                4. Paste it above to unlock AI enhancement
                """)
            else:
                st.success("✅ Claude Sonnet 4 ready!")

        # Developer info
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        **Developed by:** James Taylor  
        **Powered by:** Claude Sonnet 4  
        """)

    return enhance_markdown, api_key


def render_main_features():
    """Render the main features section using content from features.py"""
    st.markdown(get_tool_description())
    st.markdown("")  # Add spacing

    features = get_main_features()
    feature_text = ""
    for feature_key, feature_data in features.items():
        feature_text += f"**{feature_data['icon']} {feature_data['title']}:** {feature_data['description']}  \n"

    st.markdown(feature_text)


def handle_file_upload(enhance_markdown, api_key):
    """Handle file upload with enhanced processing."""
    # Get supported file extensions
    all_extensions = []
    formats = get_supported_formats()
    for category, info in formats.items():
        all_extensions.extend(info["extensions"])

    uploaded_file = st.file_uploader(
        "Select a file to convert",
        type=all_extensions,
        help="Upload any supported document format for conversion"
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
                    api_key=api_key
                )

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = uploaded_file.name

                    # Show success message
                    if enhance_markdown and api_key:
                        st.success("✨ Conversion completed successfully with Claude Sonnet 4 enhancement!")
                    else:
                        st.success("✅ Conversion completed successfully!")


def handle_folder_processing(enhance_markdown, api_key):
    """Handle batch folder processing."""
    from src.ui.folder_picker import show_folder_picker, show_output_folder_picker

    st.info(
        "🚀 **Batch Processing:** Convert all supported files in a folder to markdown format. "
        "Perfect for processing multiple documents at once."
    )

    # Input folder selection with native picker
    st.subheader("📂 Select Input Folder")
    input_folder = show_folder_picker("input")

    # Output folder selection (optional) with native picker
    st.subheader("📁 Choose Output Location")
    output_folder = show_output_folder_picker("output")

    # Process folder button
    if st.button("🚀 Process Folder", key="process_folder", type="primary"):
        if not input_folder or not os.path.isdir(input_folder):
            st.error("Please select a valid input folder using the Browse button above")
        else:
            # Process the folder
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
                    status_text.text(status)

                # Get final results
                success_count, error_count, errors = next(folder_processor)

                # Save results to session state
                st.session_state.folder_processing_results = {
                    "success_count": success_count,
                    "error_count": error_count,
                    "errors": errors,
                    "output_folder": output_folder if output_folder else os.path.join(input_folder, "markdown"),
                    "enhanced": enhance_markdown and api_key
                }

                # Show success message
                if success_count > 0:
                    st.success(f"✨ Successfully converted {success_count} files!")

                # Show error message if any
                if error_count > 0:
                    st.warning(f"Failed to convert {error_count} files. See details below.")

            except Exception as e:
                st.error(f"Error processing folder: {str(e)}")

    # Display folder processing results if available
    display_folder_results()


def display_folder_results():
    """Display folder processing results."""
    if "folder_processing_results" in st.session_state and st.session_state.folder_processing_results:
        results = st.session_state.folder_processing_results

        st.subheader("📊 Processing Results")

        # Create columns for better layout
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("✅ Successfully Converted", results['success_count'])
        with col2:
            st.metric("❌ Failed Conversions", results['error_count'])
        with col3:
            st.metric("🤖 AI Provider", "Claude Sonnet 4" if results.get('enhanced') else "Standard")

        st.markdown(f"**📁 Output Location:** {results['output_folder']}")

        # Show errors if any
        if results['error_count'] > 0:
            with st.expander("🔍 View Conversion Errors"):
                for file_name, error in results['errors'].items():
                    st.markdown(f"**{file_name}**: {error}")

        # Option to open output folder
        if st.button("📂 Open Output Folder"):
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


def display_output_section():
    """Display the markdown output section."""
    if "markdown_content" in st.session_state and st.session_state.markdown_content:
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
            help="Your converted markdown content"
        )

        # Download button
        filename = st.session_state.file_name.rsplit(".", 1)[
                       0] + ".md" if "." in st.session_state.file_name else st.session_state.file_name + ".md"

        st.download_button(
            label="📥 Download Markdown File",
            data=content,
            file_name=filename,
            mime="text/markdown",
            help="Download your converted markdown file"
        )


def main():
    """Main application function."""
    # Set up page configuration
    setup_page_config()

    # Initialize session state variables
    if "markdown_content" not in st.session_state:
        st.session_state.markdown_content = ""
    if "file_name" not in st.session_state:
        st.session_state.file_name = ""
    if "folder_processing_results" not in st.session_state:
        st.session_state.folder_processing_results = None

    # App header
    st.title("Document to Markdown Converter")
    st.subheader(get_feature_tagline())

    # Main features section
    render_main_features()

    # Set up the sidebar
    enhance_markdown, api_key = setup_enhanced_sidebar()

    # Set API key in environment if provided
    if api_key:
        os.environ["ANTHROPIC_API_KEY"] = api_key

    # Main content area - Tabs (now with About tab)
    tab1, tab2, tab3 = st.tabs(["📄 File Upload", "📁 Folder Processing", "ℹ️ About"])

    with tab1:
        handle_file_upload(enhance_markdown, api_key)

    with tab2:
        handle_folder_processing(enhance_markdown, api_key)

    with tab3:
        render_about_tab()

    # Display output section if markdown content exists
    display_output_section()


if __name__ == "__main__":
    main()