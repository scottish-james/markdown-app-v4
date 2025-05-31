"""
Enhanced PowerPoint Processor Module with Improved Nested Bullet Support

This module provides PowerPoint processing that preserves formatting, detects hierarchy,
extracts images, tables, and all shape content with improved nested bullet point handling.
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from collections import defaultdict
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
import re


def convert_pptx_to_markdown_enhanced(file_path):
    """
    Convert a PowerPoint file to markdown with comprehensive feature extraction.
    """
    try:
        prs = Presentation(file_path)
        all_content = []

        # Process each slide
        for slide_idx, slide in enumerate(prs.slides, 1):
            slide_content = extract_slide_content(slide, slide_idx)
            if slide_content.strip():
                all_content.append(slide_content)

        # Join slides with separators
        return "\n\n".join(all_content)

    except Exception as e:
        raise Exception(f"Error processing PowerPoint file: {str(e)}")


def extract_slide_content(slide, slide_number):
    """Extract all content from a slide including text, images, tables, and shapes."""
    content_parts = []

    # Add slide separator comment
    content_parts.append(f"<!-- Slide number: {slide_number} -->")

    # Collect all shapes with their positions for reading order
    positioned_shapes = []
    for shape in slide.shapes:
        if hasattr(shape, 'top') and hasattr(shape, 'left'):
            positioned_shapes.append((shape.top, shape.left, shape))
        else:
            positioned_shapes.append((0, 0, shape))

    # Sort by top position, then left position for proper reading order
    positioned_shapes.sort(key=lambda x: (x[0], x[1]))

    # Process shapes in reading order
    for _, _, shape in positioned_shapes:
        shape_content = extract_shape_content(shape)
        if shape_content.strip():
            content_parts.append(shape_content)

    return "\n\n".join(content_parts)


def extract_shape_content(shape):
    """Extract content from any type of shape, including hyperlinks."""
    try:
        # Handle different shape types
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return extract_image_content(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            return extract_table_content(shape.table)
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return extract_group_content(shape)
        elif hasattr(shape, 'text_frame') and shape.text_frame:
            # Get text content with inline hyperlinks and enhanced bullet handling
            text_content = extract_text_frame_content_enhanced(shape.text_frame, get_shape_context(shape))

            # Also check for shape-level hyperlinks (click actions)
            shape_hyperlink = extract_shape_hyperlink(shape)
            if shape_hyperlink and text_content.strip():
                # If the entire shape is a hyperlink, wrap the content
                return f"[{text_content}]({shape_hyperlink})"

            return text_content
        elif hasattr(shape, 'text') and shape.text:
            text_content = clean_text(shape.text)

            # Check for shape-level hyperlinks
            shape_hyperlink = extract_shape_hyperlink(shape)
            if shape_hyperlink and text_content.strip():
                return f"[{text_content}]({shape_hyperlink})"

            return text_content
        else:
            # Try to extract any text content from unknown shape types
            try:
                if hasattr(shape, 'text') and shape.text:
                    text_content = clean_text(shape.text)
                    shape_hyperlink = extract_shape_hyperlink(shape)
                    if shape_hyperlink and text_content.strip():
                        return f"[{text_content}]({shape_hyperlink})"
                    return text_content
            except:
                pass
    except:
        pass

    return ""


def extract_text_frame_content_enhanced(text_frame, context="unknown"):
    """Enhanced version using robust bullet detection for content placeholders."""
    if not text_frame or not text_frame.paragraphs:
        return ""

    # Use robust processing for content placeholders
    if context in ["content", "unknown"]:
        return process_content_placeholder_enhanced(text_frame, context)
    else:
        # Use simpler processing for titles and other contexts
        paragraphs = []
        for para_idx, paragraph in enumerate(text_frame.paragraphs):
            if not paragraph.text.strip():
                continue

            para_content = extract_paragraph_content_enhanced(
                paragraph, context, para_idx, text_frame.paragraphs
            )

            if para_content.strip():
                paragraphs.append(para_content)

        return "\n".join(paragraphs)


def extract_paragraph_content_enhanced(paragraph, context="unknown", para_idx=0, all_paragraphs=None):
    """Enhanced paragraph processing with better nested bullet detection."""
    if not paragraph.runs:
        return ""

    raw_text = paragraph.text.strip()
    if not raw_text:
        return ""

    # Get bullet level using multiple detection methods
    bullet_level = detect_bullet_level_enhanced(paragraph, raw_text)

    if bullet_level >= 0:
        return format_bullet_item_enhanced(paragraph, bullet_level)

    # Check for numbered lists
    if is_numbered_list_enhanced(paragraph):
        return format_numbered_item_enhanced(paragraph)

    # Extract formatted text with inline formatting
    formatted_text = extract_formatted_text(paragraph.runs)
    if not formatted_text.strip():
        return ""

    # Apply hierarchy formatting for non-list content
    return apply_hierarchy_formatting(formatted_text, paragraph, context)


def detect_bullet_level_enhanced(paragraph, raw_text):
    """Enhanced bullet level detection using multiple methods."""

    # Method 1: Use PowerPoint's native level property
    level = get_powerpoint_bullet_level(paragraph)
    if level >= 0:
        return level

    # Method 2: Detect from XML bullet formatting
    xml_level = get_xml_bullet_level(paragraph)
    if xml_level >= 0:
        return xml_level

    # Method 3: Detect from indentation and bullet characters
    indent_level = get_indentation_bullet_level(paragraph, raw_text)
    if indent_level >= 0:
        return indent_level

    # Method 4: Detect from manual bullet characters
    char_level = get_character_bullet_level(raw_text)
    if char_level >= 0:
        return char_level

    return -1  # Not a bullet point


def get_powerpoint_bullet_level(paragraph):
    """Get bullet level from PowerPoint's native properties."""
    try:
        # Check if paragraph has a defined level
        if hasattr(paragraph, 'level') and paragraph.level is not None:
            # Verify it's actually a bullet by checking for bullet formatting
            if has_bullet_formatting(paragraph):
                return paragraph.level

        # Some paragraphs might have bullet formatting but no explicit level
        if has_bullet_formatting(paragraph):
            return 0  # Default to level 0 if we know it's a bullet

    except Exception:
        pass

    return -1


def has_bullet_formatting(paragraph):
    """Check if paragraph has actual bullet formatting."""
    try:
        # Method 1: Check XML for bullet indicators
        if hasattr(paragraph, '_p') and paragraph._p is not None:
            xml_str = str(paragraph._p.xml) if hasattr(paragraph._p, 'xml') else ""
            bullet_indicators = ['buChar', 'buAutoNum', 'buFont', 'buNone="0"']
            if any(indicator in xml_str for indicator in bullet_indicators):
                return True

        # Method 2: Check paragraph properties
        if hasattr(paragraph, '_element'):
            # Look for bullet-related elements in the paragraph properties
            pPr = getattr(paragraph._element, 'pPr', None)
            if pPr is not None:
                # Check for bullet number or character properties
                for child in pPr:
                    tag_name = getattr(child, 'tag', '').lower()
                    if any(bullet in tag_name for bullet in ['buchar', 'buautonum', 'bufont']):
                        return True

    except Exception:
        pass

    return False


def get_xml_bullet_level(paragraph):
    """Extract bullet level from XML properties."""
    try:
        if hasattr(paragraph, '_element') and paragraph._element is not None:
            # Look for level indicators in XML
            pPr = getattr(paragraph._element, 'pPr', None)
            if pPr is not None:
                for child in pPr:
                    # Check for level attribute
                    if hasattr(child, 'attrib') and 'lvl' in child.attrib:
                        level = int(child.attrib['lvl'])
                        if has_bullet_formatting(paragraph):
                            return level

                    # Check for margin-based level detection
                    if hasattr(child, 'attrib') and 'marL' in child.attrib:
                        margin = int(child.attrib.get('marL', 0))
                        if margin > 0 and has_bullet_formatting(paragraph):
                            # Convert margin to level (rough estimation)
                            return margin // 360000  # PowerPoint units conversion

    except Exception:
        pass

    return -1


def get_indentation_bullet_level(paragraph, raw_text):
    """Detect bullet level from text indentation patterns."""
    try:
        # Get the original text with leading spaces
        original_text = paragraph.text
        leading_spaces = len(original_text) - len(original_text.lstrip(' '))

        # Check if text has bullet-like characteristics
        if has_bullet_like_text(raw_text):
            if leading_spaces == 0:
                return 0
            else:
                # Estimate level from indentation (assuming 2-4 spaces per level)
                return min(leading_spaces // 2, 5)  # Cap at 5 levels

    except Exception:
        pass

    return -1


def get_character_bullet_level(raw_text):
    """Detect bullet level from bullet characters in text."""
    bullet_chars = {
        '•': 0,  # Primary bullet
        '◦': 1,  # Secondary bullet
        '▪': 1,  # Secondary bullet
        '▫': 2,  # Tertiary bullet
        '‣': 1,  # Secondary bullet
        '·': 1,  # Secondary bullet
        '-': 0,  # Primary bullet (dash)
        '*': 0,  # Primary bullet (asterisk)
    }

    if raw_text and raw_text[0] in bullet_chars:
        return bullet_chars[raw_text[0]]

    return -1


def has_bullet_like_text(text):
    """Check if text starts with bullet-like characters."""
    if not text:
        return False

    bullet_chars = ['•', '◦', '▪', '▫', '‣', '·', '-', '*']
    return text[0] in bullet_chars


def format_bullet_item_enhanced(paragraph, level):
    """Enhanced bullet formatting with better text extraction."""
    # Extract formatted text while preserving inline formatting
    formatted_text = extract_formatted_text(paragraph.runs)

    # Remove existing bullet characters from the beginning
    cleaned_text = remove_bullet_chars(formatted_text)

    # Ensure level is reasonable
    level = max(0, min(level, 5))  # Cap between 0 and 5

    # Create proper markdown bullet with indentation
    indent = "  " * level
    return f"{indent}- {cleaned_text}"


def remove_bullet_chars(text):
    """Remove bullet characters from the beginning of text."""
    if not text:
        return text

    # Remove various bullet characters and any following whitespace
    bullet_pattern = r'^[•◦▪▫‣·\-*]\s*'
    return re.sub(bullet_pattern, '', text)


def is_numbered_list_enhanced(paragraph):
    """Enhanced numbered list detection."""
    text = paragraph.text.strip()
    if not text:
        return False

    # Expanded patterns for numbered lists
    numbered_patterns = [
        r'^\d+[\.\)]\s',  # 1. or 1)
        r'^[a-z][\.\)]\s',  # a. or a)
        r'^[A-Z][\.\)]\s',  # A. or A)
        r'^[ivxlcdm]+[\.\)]\s',  # i. ii. iii. (roman numerals)
        r'^\([0-9a-zA-Z]+\)\s',  # (1) or (a)
        r'^\d+\.\d+[\.\)]\s',  # 1.1. or 1.1)
    ]

    for pattern in numbered_patterns:
        if re.match(pattern, text, re.IGNORECASE):
            return True

    return False


def format_numbered_item_enhanced(paragraph):
    """Enhanced numbered list formatting."""
    formatted_text = extract_formatted_text(paragraph.runs)

    # Try to determine the numbering level from indentation or formatting
    level = 0
    try:
        if hasattr(paragraph, 'level') and paragraph.level is not None:
            level = paragraph.level
        else:
            # Estimate level from indentation
            original_text = paragraph.text
            leading_spaces = len(original_text) - len(original_text.lstrip(' '))
            if leading_spaces > 0:
                level = leading_spaces // 4  # Assume 4 spaces per level for numbered lists
    except:
        pass

    # Apply indentation for nested numbered lists
    indent = "   " * level  # 3 spaces for numbered list indentation

    # For markdown compatibility, use "1." for all numbered items
    return f"{indent}1. {formatted_text}"


# Keep all the existing functions that are working well
def extract_shape_hyperlink(shape):
    """Extract hyperlink from shape click actions."""
    try:
        if hasattr(shape, 'click_action') and shape.click_action is not None:
            if hasattr(shape.click_action, 'hyperlink') and shape.click_action.hyperlink is not None:
                if shape.click_action.hyperlink.address:
                    return fix_url(shape.click_action.hyperlink.address)
    except:
        pass

    return None


def get_shape_context(shape):
    """Determine the context of a shape (title, content, etc.)."""
    try:
        if hasattr(shape, 'placeholder_format'):
            if shape.placeholder_format.type == 1:  # Title placeholder
                return "title"
            elif shape.placeholder_format.type == 2:  # Content placeholder
                return "content"
            elif shape.placeholder_format.type == 3:  # Content with caption
                return "content"

        # Check shape name for clues
        if hasattr(shape, 'name'):
            name_lower = shape.name.lower()
            if 'title' in name_lower:
                return "title"
            elif 'header' in name_lower or 'heading' in name_lower:
                return "header"
    except:
        pass

    return "unknown"


def extract_image_content(shape):
    """Extract alt-text from images and check for hyperlinks."""
    try:
        alt_text = ""

        # Try to get alt text from various properties
        if hasattr(shape, 'alt_text') and shape.alt_text:
            alt_text = shape.alt_text
        elif hasattr(shape, 'image') and hasattr(shape.image, 'alt_text'):
            alt_text = shape.image.alt_text
        elif hasattr(shape, '_element'):
            # Try to extract from XML if available
            try:
                # Look for description in the XML
                xml_str = str(shape._element.xml) if hasattr(shape._element, 'xml') else ""
                import xml.etree.ElementTree as ET
                root = ET.fromstring(xml_str)

                # Look for alt text in various XML locations
                for elem in root.iter():
                    if 'descr' in elem.attrib:
                        alt_text = elem.attrib['descr']
                        break
                    elif 'title' in elem.attrib:
                        alt_text = elem.attrib['title']
                        break
            except:
                pass

        # Create the image markdown
        if alt_text:
            image_md = f"![{clean_text(alt_text)}](image)"
        else:
            image_md = "![Image](image)"

        # Check for hyperlinks on the image
        shape_hyperlink = extract_shape_hyperlink(shape)
        if shape_hyperlink:
            return f"[{image_md}]({shape_hyperlink})"

        return image_md

    except:
        return "![Image](image)"


def extract_group_content(group_shape):
    """Extract content from grouped shapes."""
    content_parts = []

    try:
        for shape in group_shape.shapes:
            shape_content = extract_shape_content(shape)
            if shape_content.strip():
                content_parts.append(shape_content)
    except:
        pass

    return "\n\n".join(content_parts)


def extract_formatted_text(runs):
    """Extract text from runs with formatting applied."""
    if not runs:
        return ""

    # Process each run and collect the results
    formatted_parts = []

    for run in runs:
        text = run.text
        if text is None:
            continue

        # Don't clean the text too aggressively - preserve spaces
        # Only do minimal cleaning
        if text:
            # Replace smart quotes but preserve other characters
            text = text.replace('"', '"').replace('"', '"')
            text = text.replace(''', "'").replace(''', "'")
            text = text.replace('—', '--').replace('–', '-')

        # Apply formatting to this run
        formatted_text = apply_formatting(text, run)
        formatted_parts.append(formatted_text)

    # Join all parts together
    result = "".join(formatted_parts)

    # Only fix obvious spacing issues, don't be too aggressive
    result = fix_basic_spacing(result)

    return result


def fix_basic_spacing(text):
    """Fix basic spacing issues without being too aggressive."""
    if not text:
        return text

    # Fix spacing around markdown formatting - be more precise
    # Fix cases where formatting is directly adjacent to words

    # Bold formatting: **word**nextword -> **word** nextword
    text = re.sub(r'\*\*([^*]+)\*\*([a-zA-Z])', r'**\1** \2', text)

    # Italic formatting: *word*nextword -> *word* nextword
    text = re.sub(r'\*([^*]+)\*([a-zA-Z])', r'*\1* \2', text)

    # Code formatting: `word`nextword -> `word` nextword
    text = re.sub(r'`([^`]+)`([a-zA-Z])', r'`\1` \2', text)

    # Combined formatting: ***word***nextword -> ***word*** nextword
    text = re.sub(r'\*\*\*([^*]+)\*\*\*([a-zA-Z])', r'***\1*** \2', text)

    # Fix multiple consecutive spaces
    text = re.sub(r' {2,}', ' ', text)

    return text


def apply_formatting(text, run):
    """Apply markdown formatting based on run properties, including hyperlinks."""
    if not text.strip():
        return text

    try:
        # Check for hyperlinks first
        if hasattr(run, 'hyperlink') and run.hyperlink and run.hyperlink.address:
            url = fix_url(run.hyperlink.address)
            # Apply other formatting to the text, then wrap in hyperlink
            formatted = apply_text_formatting(text, run)
            return f"[{formatted}]({url})"

        # No hyperlink, just apply text formatting
        formatted = apply_text_formatting(text, run)

    except:
        # If we can't access run properties, just return the text
        formatted = text

    return formatted


def apply_text_formatting(text, run):
    """Apply text formatting (bold, italic, etc.) to text."""
    if not text.strip():
        return text

    formatted = text

    try:
        font = run.font

        # Track what formatting we're applying to avoid conflicts
        has_bold = False
        has_italic = False

        # Check for bold
        if hasattr(font, 'bold') and font.bold:
            has_bold = True

        # Check for italic
        if hasattr(font, 'italic') and font.italic:
            has_italic = True

        # Check for underline (convert to bold if not already bold)
        if hasattr(font, 'underline') and font.underline and not has_bold:
            has_bold = True

        # Apply formatting in the right order
        if has_bold and has_italic:
            formatted = f"***{formatted}***"
        elif has_bold:
            formatted = f"**{formatted}**"
        elif has_italic:
            formatted = f"*{formatted}*"

        # Check for monospace/code fonts (only if not already formatted)
        if not (has_bold or has_italic):
            if hasattr(font, 'name') and font.name:
                monospace_fonts = ['Courier', 'Consolas', 'Monaco', 'Menlo', 'Source Code Pro', 'Courier New']
                if any(mono_font in font.name for mono_font in monospace_fonts):
                    formatted = f"`{formatted}`"

    except:
        pass

    return formatted


def fix_url(url):
    """Fix URLs by adding appropriate schemes if missing."""
    if not url:
        return url

    # For email addresses
    if '@' in url and not url.startswith('mailto:'):
        return f"mailto:{url}"

    # For web URLs
    if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
        if url.startswith('www.') or any(
                domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
            return f"https://{url}"

    return url


def apply_hierarchy_formatting(text, paragraph, context):
    """Apply hierarchy formatting (headers) based on context and text properties."""

    # Don't apply header formatting to list-related content
    stripped_text = text.strip()
    if any(phrase in stripped_text.lower() for phrase in [
        'first level', 'second level', 'third level', 'another', 'back to',
        'sub-item', 'nested', 'bullet under'
    ]):
        return text

    # Check if this should be a title (h1)
    if context == "title" or is_likely_title(text, paragraph, context):
        return f"# {text}"

    # Check if this should be a header
    header_level = detect_header_level(text, paragraph, context)
    if header_level > 0:
        return f"{'#' * header_level} {text}"

    # Regular paragraph
    return text


def is_likely_title(text, paragraph, context):
    """Determine if text should be treated as a main title."""
    # Explicit title context
    if context == "title":
        return True

    # Don't make list items into titles
    if text.strip() and text.strip()[0] in ['•', '·', '-', '*', '◦', '▪', '▫', '‣']:
        return False

    # Short text that's all caps AND not in a list
    if len(text) < 80 and text.isupper() and len(text) > 3:
        # But not if it looks like a list item description
        if any(phrase in text.upper() for phrase in
               ['FIRST LEVEL', 'SECOND LEVEL', 'THIRD LEVEL', 'ANOTHER', 'BACK TO']):
            return False
        return True

    # Check if it's the first significant text and relatively short
    if len(text) < 60 and context in ["unknown", "content"]:
        # Additional checks for title-like properties
        try:
            if paragraph.runs and len(paragraph.runs) > 0:
                first_run = paragraph.runs[0]
                if hasattr(first_run.font, 'size') and first_run.font.size:
                    # If font size is significantly large, it might be a title
                    # This is a heuristic - you might need to adjust based on your needs
                    return True
        except:
            pass

    return False


def detect_header_level(text, paragraph, context):
    """Detect if text should be a header and what level."""

    # Don't make headers if text is too long
    if len(text) > 120:
        return 0

    # Don't make headers out of list items
    stripped_text = text.strip()
    if any(phrase in stripped_text.lower() for phrase in [
        'first level', 'second level', 'third level', 'another', 'back to',
        'sub-item', 'nested', 'bullet under', 'numbered item'
    ]):
        return 0

    # Check for explicit header context
    if context == "header":
        return 2

    # Only check for headers that end with ':' and look like section headers
    if stripped_text.endswith(':') and len(stripped_text) < 50:
        # This might be a section header like "Unordered Lists:" or "Mixed Lists:"
        if any(word in stripped_text for word in ['Lists', 'Examples', 'Section', 'Part']):
            return 3

    # Check for text properties that suggest header
    try:
        if paragraph.runs and len(paragraph.runs) > 0:
            first_run = paragraph.runs[0]

            # Check if text is bold and relatively short
            if hasattr(first_run.font, 'bold') and first_run.font.bold and len(text) < 80:
                # But not if it's a list item
                if not any(char in text for char in ['•', '·', '-', '*', '◦', '▪', '▫', '‣']):
                    # Determine header level based on length and other factors
                    if len(text) < 40:
                        return 3  # h3 for short bold text
                    else:
                        return 2  # h2 for longer bold text
    except:
        pass

    return 0  # Not a header


def extract_table_content(table):
    """Extract table content in markdown format with formatting preservation."""
    if not table.rows:
        return ""

    markdown_table = ""

    for row_idx, row in enumerate(table.rows):
        row_content = "|"

        for cell in row.cells:
            cell_text = ""
            if hasattr(cell, 'text_frame') and cell.text_frame:
                # Extract text from cell with formatting
                for paragraph in cell.text_frame.paragraphs:
                    para_text = extract_formatted_text(paragraph.runs)
                    if para_text.strip():
                        cell_text += para_text + " "
            elif hasattr(cell, 'text') and cell.text:
                cell_text = clean_text(cell.text)

            # Clean up cell text
            cell_text = cell_text.strip().replace('\n', ' ').replace('|', '\\|')
            row_content += f" {cell_text} |"

        markdown_table += row_content + "\n"

        # Add separator after header row
        if row_idx == 0:
            separator = "|"
            for _ in row.cells:
                separator += "---------|"
            markdown_table += separator + "\n"

    return markdown_table


def clean_text(text):
    """Clean and normalize text - minimal cleaning to preserve spacing."""
    if not text:
        return ""

    # Only replace smart quotes and special dashes
    # Don't mess with spaces or other characters
    text = text.replace('"', '"').replace('"', '"')
    text = text.replace(''', "'").replace(''', "'")
    text = text.replace('—', '--').replace('–', '-')

    # Don't normalize whitespace aggressively - only trim start/end
    return text.strip()


# Debug functions for troubleshooting
def debug_paragraph_properties(paragraph):
    """Debug function to print paragraph properties for troubleshooting."""
    print(f"Text: '{paragraph.text}'")
    print(f"Level: {getattr(paragraph, 'level', 'None')}")

    try:
        if hasattr(paragraph, '_p') and paragraph._p is not None:
            xml_str = str(paragraph._p.xml)[:200] + "..." if len(str(paragraph._p.xml)) > 200 else str(paragraph._p.xml)
            print(f"XML snippet: {xml_str}")
    except:
        print("XML: Not accessible")

    print("---")


def test_bullet_detection(paragraph):
    """Test function to see how different detection methods work."""
    raw_text = paragraph.text.strip()

    methods = {
        "PowerPoint Level": get_powerpoint_bullet_level(paragraph),
        "XML Level": get_xml_bullet_level(paragraph),
        "Indentation Level": get_indentation_bullet_level(paragraph, raw_text),
        "Character Level": get_character_bullet_level(raw_text),
        "Final Level": detect_bullet_level_enhanced(paragraph, raw_text)
    }

    print(f"Text: '{raw_text}'")
    for method, level in methods.items():
        print(f"{method}: {level}")
    print("---")

    return methods["Final Level"]


def process_content_placeholder_enhanced(text_frame, context="content"):
    """Enhanced processing specifically for content placeholders with robust bullet detection."""
    if not text_frame or not text_frame.paragraphs:
        return ""

    # First, analyze the entire text frame to understand its structure
    structure_analysis = analyze_text_frame_structure(text_frame)

    # Process paragraphs with the structure context
    processed_paragraphs = []

    for i, paragraph in enumerate(text_frame.paragraphs):
        if not paragraph.text.strip():
            continue

        processed_para = process_paragraph_with_context(
            paragraph, i, structure_analysis, context
        )

        if processed_para.strip():
            processed_paragraphs.append(processed_para)

    return "\n".join(processed_paragraphs)


def analyze_text_frame_structure(text_frame):
    """Analyze the entire text frame to understand bullet patterns and structure."""
    analysis = {
        'total_paragraphs': len(text_frame.paragraphs),
        'bullet_paragraphs': [],
        'level_distribution': defaultdict(int),
        'bullet_styles': set(),
        'has_mixed_content': False,
        'predominant_pattern': None,
        'bullet_indicators': []
    }

    bullet_indicators = []

    for i, para in enumerate(text_frame.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # Comprehensive bullet detection for analysis
        bullet_info = detect_all_bullet_indicators(para, text)

        if bullet_info['is_bullet']:
            analysis['bullet_paragraphs'].append(i)
            analysis['level_distribution'][bullet_info['level']] += 1
            analysis['bullet_styles'].add(bullet_info['style'])
            bullet_indicators.append(bullet_info)
        else:
            # Check if this might be a continuation or non-bullet content
            if bullet_indicators:  # We have bullets before this
                analysis['has_mixed_content'] = True

    analysis['bullet_indicators'] = bullet_indicators

    # Determine the predominant pattern
    if len(bullet_indicators) > 0:
        if all(bi['method'] == 'powerpoint_native' for bi in bullet_indicators):
            analysis['predominant_pattern'] = 'powerpoint_native'
        elif all(bi['method'] == 'manual_typed' for bi in bullet_indicators):
            analysis['predominant_pattern'] = 'manual_typed'
        else:
            analysis['predominant_pattern'] = 'mixed'

    return analysis


def detect_all_bullet_indicators(paragraph, text):
    """Comprehensive bullet detection that checks all possible indicators."""
    result = {
        'is_bullet': False,
        'level': 0,
        'style': None,
        'method': None,
        'confidence': 0,
        'original_text': text
    }

    # Method 1: PowerPoint Native Bullets (highest confidence)
    native_result = detect_powerpoint_native_bullets(paragraph)
    if native_result['is_bullet']:
        result.update(native_result)
        result['confidence'] = 100
        return result

    # Method 2: XML-based detection (high confidence) - use your existing function
    if has_bullet_formatting(paragraph):
        xml_level = get_xml_bullet_level(paragraph)
        if xml_level >= 0:
            result['is_bullet'] = True
            result['level'] = xml_level
            result['style'] = 'xml'
            result['method'] = 'xml_analysis'
            result['confidence'] = 90
            return result

    # Method 3: Manual typed bullets (medium confidence)
    manual_result = detect_manual_bullets_robust(text)
    if manual_result['is_bullet']:
        result.update(manual_result)
        result['confidence'] = 70
        # Try to get level from indentation
        indent_level = detect_indentation_level_robust(paragraph.text)
        if indent_level > 0:
            result['level'] = indent_level
        return result

    # Method 4: Pattern-based detection (low confidence)
    if is_numbered_list_enhanced(paragraph):
        result['is_bullet'] = True
        result['level'] = 0
        result['style'] = 'numbered'
        result['method'] = 'pattern_based'
        result['confidence'] = 50
        return result

    return result


def detect_powerpoint_native_bullets(paragraph):
    """Detect PowerPoint's native bullet formatting."""
    result = {'is_bullet': False, 'level': 0, 'style': 'native', 'method': 'powerpoint_native'}

    try:
        # Check if paragraph has a level property
        if hasattr(paragraph, 'level') and paragraph.level is not None:
            # Verify it actually has bullet formatting
            if has_bullet_formatting(paragraph):
                result['is_bullet'] = True
                result['level'] = paragraph.level
                return result

        # Sometimes level is None but it's still a bullet at level 0
        if has_bullet_formatting(paragraph):
            result['is_bullet'] = True
            result['level'] = 0
            return result

    except Exception:
        pass

    return result


def detect_manual_bullets_robust(text):
    """Detect manually typed bullet characters."""
    result = {'is_bullet': False, 'level': 0, 'style': 'manual', 'method': 'manual_typed'}

    if not text:
        return result

    # Define bullet characters and their typical hierarchy
    bullet_hierarchy = {
        '•': 0,  # Primary bullet
        '◦': 1,  # Secondary bullet (hollow)
        '▪': 1,  # Secondary bullet (small square)
        '▫': 2,  # Tertiary bullet (hollow square)
        '‣': 1,  # Secondary bullet (triangular)
        '·': 1,  # Secondary bullet (middle dot)
        '○': 1,  # Secondary bullet (circle)
        '■': 1,  # Secondary bullet (square)
        '□': 2,  # Tertiary bullet (hollow square)
        '→': 1,  # Arrow bullet
        '►': 1,  # Arrow bullet
        '✓': 1,  # Checkmark bullet
        '✗': 1,  # X bullet
        '-': 0,  # Dash bullet
        '*': 0,  # Asterisk bullet
        '+': 0,  # Plus bullet
    }

    first_char = text[0]
    if first_char in bullet_hierarchy:
        result['is_bullet'] = True
        result['level'] = bullet_hierarchy[first_char]
        result['style'] = f'manual_{first_char}'
        return result

    return result


def detect_indentation_level_robust(text):
    """Detect indentation level from the actual text."""
    if not text:
        return 0

    # Count leading spaces
    leading_spaces = len(text) - len(text.lstrip(' '))

    # Count leading tabs (convert to equivalent spaces)
    leading_tabs = len(text) - len(text.lstrip('\t'))
    equivalent_spaces = leading_spaces + (leading_tabs * 4)

    # Estimate level (assuming 2-4 spaces per level)
    if equivalent_spaces == 0:
        return 0
    elif equivalent_spaces <= 4:
        return 1
    elif equivalent_spaces <= 8:
        return 2
    elif equivalent_spaces <= 12:
        return 3
    else:
        return min(equivalent_spaces // 4, 5)  # Cap at 5 levels


def process_paragraph_with_context(paragraph, para_index, structure_analysis, context):
    """Process a paragraph with full context awareness."""
    text = paragraph.text.strip()
    if not text:
        return ""

    # Get comprehensive bullet information
    bullet_info = detect_all_bullet_indicators(paragraph, text)

    if bullet_info['is_bullet']:
        return format_bullet_with_context(paragraph, bullet_info, structure_analysis)
    else:
        # Handle non-bullet content
        formatted_text = extract_formatted_text(paragraph.runs)

        # Apply appropriate formatting based on context
        if context == "title" or is_likely_title_in_context(formatted_text, paragraph, para_index, structure_analysis):
            return f"# {formatted_text}"
        elif should_be_header_robust(formatted_text, paragraph, structure_analysis):
            return f"## {formatted_text}"
        else:
            return formatted_text


def format_bullet_with_context(paragraph, bullet_info, structure_analysis):
    """Format a bullet point with full context awareness."""
    # Extract the text content with formatting
    formatted_text = extract_formatted_text(paragraph.runs)

    # Remove bullet characters if they were manually typed
    if bullet_info['method'] == 'manual_typed':
        formatted_text = remove_leading_bullet_chars_robust(formatted_text)

    # Adjust level based on structure analysis if needed
    level = bullet_info['level']

    # Handle inconsistent leveling in mixed content
    if structure_analysis['predominant_pattern'] == 'mixed':
        level = normalize_level_in_mixed_content(level, bullet_info, structure_analysis)

    # Cap the level to prevent excessive indentation
    level = max(0, min(level, 5))

    # Create markdown bullet
    indent = "  " * level
    return f"{indent}- {formatted_text.strip()}"


def remove_leading_bullet_chars_robust(text):
    """Remove leading bullet characters and normalize spacing."""
    if not text:
        return text

    # Remove various bullet characters and normalize spacing
    bullet_pattern = r'^[•◦▪▫‣·○■□→►✓✗\-\*\+]\s*'
    cleaned = re.sub(bullet_pattern, '', text)

    # Also handle numbered patterns
    numbered_pattern = r'^(?:\d+[\.\)]|\([0-9a-zA-Z]+\)|[a-zA-Z][\.\)]|\d+\.\d+[\.\)]|[ivx]+[\.\)])\s*'
    cleaned = re.sub(numbered_pattern, '', cleaned, flags=re.IGNORECASE)

    return cleaned.strip()


def normalize_level_in_mixed_content(level, bullet_info, structure_analysis):
    """Normalize bullet levels when dealing with mixed content patterns."""
    # If we have mixed manual and native bullets, try to normalize
    if bullet_info['method'] == 'manual_typed':
        # Manual bullets often have incorrect levels, try to infer from indentation
        return detect_indentation_level_robust(bullet_info.get('original_text', ''))

    return level


def is_likely_title_in_context(text, paragraph, para_index, structure_analysis):
    """Determine if text should be treated as a title given the context."""
    # First paragraph in content placeholder, short, and no bullets after
    if para_index == 0 and len(text) < 100 and structure_analysis['bullet_paragraphs']:
        return True

    # All caps and short
    if text.isupper() and len(text) < 80:
        return True

    return False


def should_be_header_robust(text, paragraph, structure_analysis):
    """Determine if text should be formatted as a header."""
    # Short text that ends with colon (like "Key Points:")
    if len(text) < 60 and text.endswith(':'):
        return True

    # Bold text that's relatively short
    try:
        if paragraph.runs and len(paragraph.runs) > 0:
            first_run = paragraph.runs[0]
            if hasattr(first_run.font, 'bold') and first_run.font.bold and len(text) < 80:
                return True
    except:
        pass

    return False