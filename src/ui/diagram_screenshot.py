"""
Debug Screenshot UI with better slide mapping and debug output
Save as: src/ui/diagram_screenshot_ui.py
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
    DEBUG VERSION - Shows slide mapping process
    """
    st.markdown("---")
    st.subheader("🐛 DEBUG: Diagram Screenshot Analysis")

    # Step 1: Check session state
    st.markdown("**Step 1: Session State Check**")
    has_content = 'markdown_content' in st.session_state and st.session_state.markdown_content
    has_filename = 'file_name' in st.session_state and st.session_state.file_name
    has_file_data = 'uploaded_file_data' in st.session_state

    st.write(f"• Has markdown content: {'✅' if has_content else '❌'}")
    st.write(f"• Has filename: {'✅' if has_filename else '❌'}")
    st.write(f"• Has file data: {'✅' if has_file_data else '❌'}")

    if has_file_data:
        file_data_size = len(st.session_state.uploaded_file_data)
        st.write(f"• File data size: {file_data_size / (1024 * 1024):.1f} MB")

    if not (has_content and has_filename and has_file_data):
        st.error("❌ Missing required data - upload and convert a PowerPoint file first")
        return

    # Step 2: Check file type
    st.markdown("**Step 2: File Type Check**")
    filename = st.session_state.file_name
    file_ext = filename.split('.')[-1].lower() if '.' in filename else ''
    is_powerpoint = file_ext in ['pptx', 'ppt']

    st.write(f"• Filename: {filename}")
    st.write(f"• Extension: .{file_ext}")
    st.write(f"• Is PowerPoint: {'✅' if is_powerpoint else '❌'}")

    if not is_powerpoint:
        st.error("❌ Not a PowerPoint file")
        return

    # Step 3: Parse v19 analysis
    st.markdown("**Step 3: v19 Analysis Search**")
    markdown_content = st.session_state.markdown_content

    high_prob_slides = _parse_v19_diagram_analysis(markdown_content)

    if not high_prob_slides:
        st.error("❌ No slides >80% probability found")

        # Show all slides for debugging
        all_slides = _parse_all_slides_from_markdown(markdown_content)
        if all_slides:
            st.write("**All slides found in analysis:**")
            for slide in all_slides:
                st.write(f"  • Slide {slide['slide_number']}: {slide['probability']}%")
        else:
            st.write("❌ No v19 DIAGRAM ANALYSIS section found")
        return

    st.write(f"✅ Found {len(high_prob_slides)} slides >80%:")
    for slide in high_prob_slides:
        st.write(f"  • Slide {slide['slide_number']}: {slide['probability']}%")

    # Step 4: Screenshot button with debug
    st.markdown("**Step 4: Screenshot with Debug Info**")
    slide_numbers = [s['slide_number'] for s in high_prob_slides]
    slide_text = ", ".join(map(str, slide_numbers))

    if st.button(f"🧪 DEBUG SCREENSHOT: {slide_text}", type="primary"):
        _execute_debug_screenshots(slide_numbers)


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


def _parse_all_slides_from_markdown(markdown_content: str) -> List[Dict]:
    """Parse all slides from markdown for debugging."""
    all_slides = []

    if 'DIAGRAM ANALYSIS' not in markdown_content:
        return all_slides

    analysis_pattern = r'# DIAGRAM ANALYSIS.*?\*\*Slides with potential diagrams:\*\*(.*?)(?=\n#|\Z)'
    analysis_match = re.search(analysis_pattern, markdown_content, re.DOTALL | re.IGNORECASE)

    if analysis_match:
        analysis_section = analysis_match.group(1)
        slide_pattern = r'- \*\*Slide (\d+)\*\*: (\d+)% probability'

        for match in re.finditer(slide_pattern, analysis_section):
            slide_num = int(match.group(1))
            probability = int(match.group(2))

            all_slides.append({
                'slide_number': slide_num,
                'probability': probability
            })

    return sorted(all_slides, key=lambda x: x['probability'], reverse=True)


def _execute_debug_screenshots(slide_numbers: List[int]):
    """Execute screenshots with full debug output."""
    st.markdown("---")
    st.subheader("🔧 Debug Screenshot Process")

    try:
        # Setup
        filename = st.session_state.file_name
        file_ext = filename.split('.')[-1].lower() if '.' in filename else 'pptx'
        file_data = st.session_state.uploaded_file_data

        st.write(f"📁 File: {filename}")
        st.write(f"📊 Size: {len(file_data) / (1024 * 1024):.1f} MB")
        st.write(f"🎯 Target slides: {slide_numbers}")

        # Create temp file
        with tempfile.NamedTemporaryFile(suffix=f".{file_ext}", delete=False) as temp_pptx:
            temp_pptx.write(file_data)
            temp_pptx_path = temp_pptx.name

        st.write(f"💾 Temp file: {temp_pptx_path}")

        # Test LibreOffice
        from src.processors.diagram_screenshot_processor import test_diagram_screenshot_capability
        available, status = test_diagram_screenshot_capability()
        st.write(f"🔧 LibreOffice: {status}")

        if not available:
            st.error("❌ LibreOffice not available")
            return

        # Screenshot with debug
        with tempfile.TemporaryDirectory() as output_dir:
            st.write(f"📂 Output: {output_dir}")

            from src.processors.diagram_screenshot_processor import DiagramScreenshotProcessor
            processor = DiagramScreenshotProcessor()

            # Capture debug output
            debug_output = StringIO()

            with st.spinner("📸 Generating screenshots with debug..."):
                # Redirect print statements to capture debug info
                with redirect_stdout(debug_output):
                    results = processor._screenshot_specific_slides(
                        temp_pptx_path,
                        slide_numbers,
                        output_dir,
                        filename.rsplit('.', 1)[0],
                        debug_mode=True  # Enable debug mode
                    )

            # Show debug output
            debug_text = debug_output.getvalue()
            if debug_text:
                st.markdown("**🔍 LibreOffice Debug Output:**")
                st.code(debug_text)

            # Show results
            st.markdown("**📊 Screenshot Results:**")

            if results:
                st.success(f"✅ Generated {len(results)} screenshots")

                # Show each result
                for slide_num in slide_numbers:
                    if slide_num in results:
                        file_path = results[slide_num]
                        st.success(f"✅ Slide {slide_num}: {os.path.basename(file_path)}")

                        # Show thumbnail
                        if os.path.exists(file_path):
                            col1, col2 = st.columns([1, 2])

                            with col1:
                                st.image(file_path, caption=f"Slide {slide_num}", width=200)

                            with col2:
                                # File info
                                file_size = os.path.getsize(file_path) / 1024
                                st.write(f"**File:** {os.path.basename(file_path)}")
                                st.write(f"**Size:** {file_size:.1f} KB")

                                # Download button
                                with open(file_path, "rb") as f:
                                    image_data = f.read()

                                st.download_button(
                                    f"📥 Download Slide {slide_num}",
                                    data=image_data,
                                    file_name=os.path.basename(file_path),
                                    mime="image/png",
                                    key=f"download_debug_{slide_num}"
                                )
                    else:
                        st.error(f"❌ Slide {slide_num}: No screenshot generated")
            else:
                st.error("❌ No screenshots were generated")

        # Cleanup
        os.unlink(temp_pptx_path)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        with st.expander("🐛 Full Error"):
            st.code(traceback.format_exc())


def store_uploaded_file_data(uploaded_file):
    """Store PowerPoint file data."""
    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower() if '.' in uploaded_file.name else ''
        if file_ext in ['pptx', 'ppt']:
            st.session_state.uploaded_file_data = uploaded_file.getbuffer()


def clear_screenshot_data():
    """Clear screenshot data."""
    if 'uploaded_file_data' in st.session_state:
        del st.session_state.uploaded_file_data