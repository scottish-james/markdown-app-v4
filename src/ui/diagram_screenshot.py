"""
Fixed Screenshot UI - Uses new fallback method for reliable slide export
Save as: src/ui/diagram_screenshot.py
"""

import streamlit as st
import re
import tempfile
import os
import sys
from io import StringIO
from contextlib import redirect_stdout
from typing import List, Dict


def render_diagram_screenshot_section():
    """
    FIXED VERSION - Uses new fallback method for reliable slide screenshots
    """
    st.markdown("---")
    st.subheader("📸 Diagram Screenshot Generator")

    # Step 1: Check session state
    st.markdown("**Session State Check**")
    has_content = 'markdown_content' in st.session_state and st.session_state.markdown_content
    has_filename = 'file_name' in st.session_state and st.session_state.file_name
    has_file_data = 'uploaded_file_data' in st.session_state

    status_col1, status_col2, status_col3 = st.columns(3)
    with status_col1:
        st.write("✅ Content" if has_content else "❌ Content")
    with status_col2:
        st.write("✅ Filename" if has_filename else "❌ Filename")
    with status_col3:
        st.write("✅ File Data" if has_file_data else "❌ File Data")

    if not (has_content and has_filename and has_file_data):
        st.warning("⚠️ Upload and convert a PowerPoint file first to enable screenshots")
        return

    # Step 2: Check file type
    filename = st.session_state.file_name
    file_ext = filename.split('.')[-1].lower() if '.' in filename else ''
    is_powerpoint = file_ext in ['pptx', 'ppt']

    if not is_powerpoint:
        st.error("❌ Screenshots only work with PowerPoint files (.pptx, .ppt)")
        return

    st.success(f"✅ PowerPoint file ready: {filename}")

    # Step 3: Parse v19 analysis for high-probability slides
    markdown_content = st.session_state.markdown_content
    high_prob_slides = _parse_v19_diagram_analysis(markdown_content)

    if not high_prob_slides:
        st.warning("⚠️ No slides with >80% diagram probability found")

        # Show manual slide selection option
        with st.expander("🔧 Manual Slide Selection", expanded=False):
            st.markdown("**Select slides manually if automatic detection didn't work:**")

            # Get total slide count from markdown
            total_slides = _count_total_slides(markdown_content)
            if total_slides > 0:
                st.write(f"Total slides in presentation: {total_slides}")

                # Manual slide number input
                manual_slides = st.text_input(
                    "Enter slide numbers (comma-separated):",
                    placeholder="e.g., 1, 3, 5",
                    help="Enter the slide numbers you want to screenshot"
                )

                if manual_slides and st.button("📸 Screenshot Selected Slides", type="secondary"):
                    try:
                        slide_numbers = [int(x.strip()) for x in manual_slides.split(',')]
                        slide_numbers = [s for s in slide_numbers if 1 <= s <= total_slides]
                        if slide_numbers:
                            _execute_improved_screenshots(slide_numbers)
                        else:
                            st.error("❌ Invalid slide numbers")
                    except ValueError:
                        st.error("❌ Please enter valid numbers separated by commas")
            else:
                st.write("Could not determine total slide count from markdown")
        return

    # Show detected high-probability slides
    st.success(f"🎯 Found {len(high_prob_slides)} slides with >80% diagram probability:")

    for slide in high_prob_slides:
        st.write(f"• **Slide {slide['slide_number']}**: {slide['probability']}% probability")

    # Main screenshot button
    slide_numbers = [s['slide_number'] for s in high_prob_slides]
    slide_text = ", ".join(map(str, slide_numbers))

    if st.button(f"📸 Screenshot Diagram Slides: {slide_text}", type="primary"):
        _execute_improved_screenshots(slide_numbers)


def _parse_v19_diagram_analysis(markdown_content: str) -> List[Dict]:
    """Parse v19 analysis for slides >80%."""
    high_prob_slides = []

    # Find analysis section
    analysis_pattern = r'# DIAGRAM ANALYSIS.*?\*\*Slides with potential diagrams:\*\*(.*?)(?=\n#|\Z)'
    analysis_match = re.search(analysis_pattern, markdown_content, re.DOTALL | re.IGNORECASE)

    if not analysis_match:
        return high_prob_slides

    analysis_section = analysis_match.group(1)

    # Parse slide lines
    slide_pattern = r'- \*\*Slide (\d+)\*\*: (\d+)% probability'

    for match in re.finditer(slide_pattern, analysis_section):
        slide_num = int(match.group(1))
        probability = int(match.group(2))

        if probability > 80:
            high_prob_slides.append({
                'slide_number': slide_num,
                'probability': probability
            })

    return sorted(high_prob_slides, key=lambda x: x['probability'], reverse=True)


def _count_total_slides(markdown_content: str) -> int:
    """Count total slides from slide markers in markdown."""
    slide_markers = re.findall(r'<!-- Slide (\d+) -->', markdown_content)
    if slide_markers:
        return max(int(marker) for marker in slide_markers)
    return 0


def _execute_improved_screenshots(slide_numbers: List[int]):
    """Execute screenshots using the new improved method with fallback."""
    st.markdown("---")
    st.subheader("📸 Screenshot Generation Process")

    try:
        # Setup
        filename = st.session_state.file_name
        file_ext = filename.split('.')[-1].lower() if '.' in filename else 'pptx'
        file_data = st.session_state.uploaded_file_data

        st.write(f"📁 **File:** {filename}")
        st.write(f"📊 **Size:** {len(file_data) / (1024 * 1024):.1f} MB")
        st.write(f"🎯 **Target slides:** {slide_numbers}")

        # Create temp file
        with tempfile.NamedTemporaryFile(suffix=f".{file_ext}", delete=False) as temp_pptx:
            temp_pptx.write(file_data)
            temp_pptx_path = temp_pptx.name

        # Test LibreOffice availability
        from src.processors.diagram_screenshot_processor import test_diagram_screenshot_capability, \
            install_poppler_instructions
        available, status = test_diagram_screenshot_capability()
        st.write(f"🔧 **LibreOffice Status:** {status}")

        if not available:
            st.error("❌ LibreOffice not available - required for screenshot generation")
            return

        # Show poppler installation tip if not available
        if "without Poppler" in status:
            with st.expander("💡 Improve Screenshot Quality"):
                st.markdown("**Install poppler-utils for better screenshot quality:**")
                instructions = install_poppler_instructions()
                st.code(instructions)
                st.info(
                    "Poppler-utils provides more reliable PDF page extraction. The tool will work without it, but with poppler you'll get better results.")

        # Progress indicator
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Screenshot with improved method
        with tempfile.TemporaryDirectory() as output_dir:
            from src.processors.diagram_screenshot_processor import DiagramScreenshotProcessor
            processor = DiagramScreenshotProcessor()

            status_text.text("📸 Generating screenshots...")
            progress_bar.progress(25)

            # Capture debug output for display
            debug_output = StringIO()

            with st.spinner("📸 Generating screenshots..."):
                # Redirect print statements to capture debug info
                with redirect_stdout(debug_output):
                    # FIXED: Use the correct method name
                    results = processor.screenshot_slides_with_all_methods(
                        temp_pptx_path,
                        slide_numbers,
                        output_dir,
                        filename.rsplit('.', 1)[0],
                        debug_mode=True  # Show debug info
                    )

            # Show debug output
            debug_text = debug_output.getvalue()
            if debug_text:
                with st.expander("🔍 LibreOffice Debug Output"):
                    st.code(debug_text)

            progress_bar.progress(75)
            status_text.text("Processing results...")

            # Clear progress indicators
            progress_bar.progress(100)
            status_text.text("✅ Screenshot generation complete!")

            # Show results
            st.markdown("---")
            st.subheader("📊 Screenshot Results")

            if results:
                success_count = len(results)
                total_count = len(slide_numbers)

                if success_count == total_count:
                    st.success(f"🎉 Successfully generated {success_count} screenshots!")
                else:
                    st.warning(f"⚠️ Generated {success_count} of {total_count} requested screenshots")

                # Display each screenshot
                for slide_num in slide_numbers:
                    if slide_num in results:
                        _display_screenshot_result(slide_num, results[slide_num])
                    else:
                        st.error(f"❌ **Slide {slide_num}**: Screenshot generation failed")

            else:
                st.error("❌ No screenshots were generated")
                st.info("💡 **Troubleshooting:**")
                st.markdown("""
                - Make sure LibreOffice is properly installed
                - Try with different slides
                - Check if the PowerPoint file is valid
                - Some slides might not be exportable due to complex formatting
                """)

        # Cleanup
        os.unlink(temp_pptx_path)
        progress_bar.empty()
        status_text.empty()

    except Exception as e:
        st.error(f"❌ **Error during screenshot generation:** {str(e)}")

        with st.expander("🐛 Debug Information"):
            import traceback
            st.code(traceback.format_exc())


def _display_screenshot_result(slide_num: int, file_path: str):
    """Display a single screenshot result with thumbnail and download."""
    if not os.path.exists(file_path):
        st.error(f"❌ Slide {slide_num}: File not found")
        return

    st.markdown(f"### ✅ Slide {slide_num}")

    col1, col2 = st.columns([1, 2])

    with col1:
        # Display thumbnail
        try:
            st.image(file_path, caption=f"Slide {slide_num}", width=250)
        except Exception as e:
            st.error(f"Error displaying image: {str(e)}")

    with col2:
        # File information
        file_size = os.path.getsize(file_path) / 1024
        st.write(f"**📄 Filename:** {os.path.basename(file_path)}")
        st.write(f"**📊 Size:** {file_size:.1f} KB")

        # Download button
        try:
            with open(file_path, "rb") as f:
                image_data = f.read()

            st.download_button(
                f"📥 Download Slide {slide_num}",
                data=image_data,
                file_name=os.path.basename(file_path),
                mime="image/png",
                key=f"download_slide_{slide_num}",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error creating download button: {str(e)}")

    st.markdown("---")


def store_uploaded_file_data(uploaded_file):
    """Store PowerPoint file data for screenshot functionality."""
    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower() if '.' in uploaded_file.name else ''
        if file_ext in ['pptx', 'ppt']:
            st.session_state.uploaded_file_data = uploaded_file.getbuffer()
            st.session_state.uploaded_file_name = uploaded_file.name


def clear_screenshot_data():
    """Clear screenshot-related session data."""
    keys_to_clear = ['uploaded_file_data', 'uploaded_file_name']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]


def debug_screenshot_info():
    """Display debug information about screenshot functionality."""
    with st.expander("🐛 Screenshot Debug Info"):
        st.markdown("**Session State:**")
        for key in ['markdown_content', 'file_name', 'uploaded_file_data']:
            has_key = key in st.session_state
            st.write(f"• {key}: {'✅' if has_key else '❌'}")

            if has_key and key == 'uploaded_file_data':
                size_mb = len(st.session_state[key]) / (1024 * 1024)
                st.write(f"  Size: {size_mb:.2f} MB")

        # LibreOffice status
        st.markdown("**LibreOffice Status:**")
        from src.processors.diagram_screenshot_processor import test_diagram_screenshot_capability
        available, status = test_diagram_screenshot_capability()
        st.write(f"• Status: {status}")
        st.write(f"• Available: {'✅' if available else '❌'}")