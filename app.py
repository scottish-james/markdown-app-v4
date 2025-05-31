"""
Integration script to add Claude Sonnet 4 support to your existing Streamlit app.

Add this to your existing app.py or create a new enhanced version.
"""

import streamlit as st
import os
from src.converters.file_converter import convert_file_to_markdown
from src.converters.url_converter import convert_url_to_markdown
from src.processors.folder_processor import process_folder
from src.ui.components import setup_sidebar, get_supported_formats, setup_page_config


# Updated sidebar function with Claude support
def setup_sidebar_with_claude():
    """Enhanced sidebar setup with Claude Sonnet 4 support."""
    with st.sidebar:
        st.header("Office to MD")
        st.write("Document Conversion Tool")

        # Supported formats in an expander
        with st.expander("Supported Formats"):
            from src.ui.components import display_supported_formats
            display_supported_formats()
            st.markdown("**🌐 Websites**")
            st.markdown("- Any URL (converts HTML to Markdown)")

        # Enhancement options
        st.subheader("Enhancement Options")
        enhance_markdown = st.checkbox("Enhance with AI", value=True,
                                       help="Use AI to improve markdown formatting")

        # Provider selection
        enhancement_provider = "claude"  # Default to Claude
        api_key = None

        if enhance_markdown:
            provider_option = st.selectbox(
                "AI Provider",
                ["Claude Sonnet 4 (Recommended)", "OpenAI GPT"],
                help="Choose which AI service to use for enhancement"
            )

            if "Claude" in provider_option:
                enhancement_provider = "claude"
                api_key = st.text_input(
                    "Anthropic API Key",
                    type="password",
                    help="Enter your Anthropic API key for Claude enhancement"
                )

                # Show Claude advantages
                st.info("🎯 Claude Sonnet 4 provides superior document structure analysis and formatting preservation!")

            else:
                enhancement_provider = "openai"
                api_key = st.text_input(
                    "OpenAI API Key",
                    type="password",
                    help="Enter your OpenAI API key for enhancement"
                )

        # Developer info
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        **Developed by:** James Taylor
        **Enhanced with:** Claude Sonnet 4 Support
        """)

    return enhance_markdown, api_key, enhancement_provider


def handle_file_upload_enhanced(enhance_markdown, api_key, enhancement_provider):
    """Enhanced file upload handler with Claude support."""
    # Get supported file extensions
    all_extensions = []
    formats = get_supported_formats()
    for category, info in formats.items():
        all_extensions.extend(info["extensions"])

    # Enhanced info message
    if enhancement_provider == "claude":
        st.info(
            "🚀 **Claude Sonnet 4 Enhancement Active!** "
            "Your documents will be processed with advanced structure analysis and formatting preservation. "
            "PDF and PowerPoint files will have their hyperlinks extracted and included in the markdown output."
        )
    else:
        st.info(
            "Special feature: PDF and PowerPoint files will have their hyperlinks extracted and included in the markdown "
            "output. However this does not work great for PDF so please avoid where possible and use original documents. "
            "WORD documents work best"
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
            with st.spinner(f"Converting to Markdown with {enhancement_provider.title()}..."):
                # Show progress bar
                progress_bar = st.progress(0)
                for percent_complete in range(100):
                    progress_bar.progress(percent_complete + 1)

                # Convert uploaded file with the selected provider
                markdown_content, error = convert_file_to_markdown(
                    uploaded_file.getbuffer(),
                    uploaded_file.name,
                    enhance=enhance_markdown,
                    api_key=api_key,
                    enhancement_provider=enhancement_provider
                )

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = uploaded_file.name

                    # Show success message with provider info
                    if enhance_markdown and api_key:
                        st.success(
                            f"✨ Conversion completed successfully with {enhancement_provider.title()} enhancement!")
                    else:
                        st.success("Conversion completed successfully!")


def main_enhanced():
    """Enhanced main application function with Claude support."""

    # Set up page configuration
    setup_page_config()

    # Initialize session state variables
    if "markdown_content" not in st.session_state:
        st.session_state.markdown_content = ""
    if "file_name" not in st.session_state:
        st.session_state.file_name = ""
    if "url_title" not in st.session_state:
        st.session_state.url_title = ""
    if "folder_processing_results" not in st.session_state:
        st.session_state.folder_processing_results = None

    # App header with Claude branding
    st.title("🚀 Office to Markdown Converter")
    st.markdown(
        "**Enhanced with Claude Sonnet 4** - Convert your documents or websites to clean, structured Markdown with advanced AI enhancement")

    # Set up the enhanced sidebar
    enhance_markdown, api_key, enhancement_provider = setup_sidebar_with_claude()

    # Set API key in environment if provided
    if api_key:
        if enhancement_provider == "claude":
            os.environ["ANTHROPIC_API_KEY"] = api_key
        else:
            os.environ["OPENAI_API_KEY"] = api_key

    # Main content area - Tabs
    tab1, tab2, tab3 = st.tabs(["📄 File Upload", "🌐 Website URL", "📁 Folder Processing"])

    with tab1:
        handle_file_upload_enhanced(enhance_markdown, api_key, enhancement_provider)

    with tab2:
        handle_url_conversion_enhanced(enhance_markdown, api_key, enhancement_provider)

    with tab3:
        handle_folder_processing_enhanced(enhance_markdown, api_key, enhancement_provider)

    # Display output section if markdown content exists
    display_output_section_enhanced()


def handle_url_conversion_enhanced(enhance_markdown, api_key, enhancement_provider):
    """Enhanced website URL conversion logic with Claude support."""
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
            with st.spinner(f"Fetching website and converting to Markdown with {enhancement_provider.title()}..."):
                # Show progress bar
                progress_bar = st.progress(0)
                for percent_complete in range(100):
                    progress_bar.progress(percent_complete + 1)

                # Convert website URL - Note: URL converter may need updating for provider selection
                try:
                    if enhancement_provider == "claude":
                        # You'll need to update url_converter.py to support Claude
                        markdown_content, error = convert_url_to_markdown(
                            website_url,
                            enhance=enhance_markdown,
                            api_key=api_key
                        )
                    else:
                        markdown_content, error = convert_url_to_markdown(
                            website_url,
                            enhance=enhance_markdown,
                            api_key=api_key
                        )
                except Exception as e:
                    markdown_content, error = "", str(e)

                if error:
                    st.error(f"Error during conversion: {error}")
                else:
                    st.session_state.markdown_content = markdown_content
                    st.session_state.file_name = f"{st.session_state.url_title}.md"

                    if enhance_markdown and api_key:
                        st.success(
                            f"✨ Website conversion completed successfully with {enhancement_provider.title()} enhancement!")
                    else:
                        st.success("Website conversion completed successfully!")


def handle_folder_processing_enhanced(enhance_markdown, api_key, enhancement_provider):
    """Enhanced batch folder processing logic with Claude support."""
    st.info(
        f"🚀 Convert all supported files in a folder to markdown format using {enhancement_provider.title()} enhancement. "
        "Each file will be processed and a .md file will be created with advanced formatting."
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

    # Enhancement provider info
    if enhancement_provider == "claude":
        st.info(
            "🎯 **Claude Sonnet 4 will be used for batch processing** - Expect superior document structure analysis and formatting!")

    # Process folder button
    if st.button("Process Folder", key="process_folder"):
        if not input_folder or not os.path.isdir(input_folder):
            st.error("Please enter a valid folder path")
        else:
            # Process the folder with enhanced processor
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # Create enhanced folder processor (you may need to update folder_processor.py)
                folder_processor = process_folder(
                    input_folder,
                    output_folder,
                    enhance=enhance_markdown,
                    api_key=api_key,
                    enhancement_provider=enhancement_provider  # Add this parameter
                )

                # Process files with progress updates
                for progress, status in folder_processor:
                    progress_bar.progress(min(1.0, progress))
                    status_text.text(f"{status} (using {enhancement_provider.title()})")

                # Get final results
                success_count, error_count, errors = next(folder_processor)

                # Save results to session state
                st.session_state.folder_processing_results = {
                    "success_count": success_count,
                    "error_count": error_count,
                    "errors": errors,
                    "output_folder": output_folder if output_folder else os.path.join(input_folder, "markdown"),
                    "enhancement_provider": enhancement_provider
                }

                # Show success message
                if success_count > 0:
                    st.success(f"✨ Successfully converted {success_count} files using {enhancement_provider.title()}!")

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
            provider = results.get('enhancement_provider', 'Unknown')
            st.metric("🤖 AI Provider", provider.title())

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
            help="Download your enhanced markdown file"
        )

        # Option to copy to clipboard (using streamlit-extras if available)
        try:
            import streamlit_extras
            if st.button("📋 Copy to Clipboard"):
                st.write("Content copied! (You may need to manually copy from the text area above)")
        except ImportError:
            pass


# Installation requirements note
def show_installation_requirements():
    """Show installation requirements for Claude integration."""
    st.markdown("""
    ## 🔧 Installation Requirements for Claude Integration

    To use Claude Sonnet 4 enhancement, you need to install the Anthropic package:

    ```bash
    pip install anthropic
    ```

    You'll also need an Anthropic API key. Get one at: https://console.anthropic.com/
    """)


if __name__ == "__main__":
    main_enhanced()