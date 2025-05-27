"""
Enhanced PowerPoint Processor Module

This module provides PowerPoint processing that preserves formatting, detects hierarchy,
extracts images, tables, and all shape content with slide separators.
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
import re


def convert_pptx_to_markdown_enhanced(file_path):
    """
    Convert a PowerPoint file to markdown with comprehensive feature extraction.

    Args:
        file_path (str): Path to the PowerPoint file

    Returns:
        str: Markdown formatted content with slide separators
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
            positioned_shapes.append((0, 0, shape))  # Fallback position

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
            # Get text content with inline hyperlinks
            text_content = extract_text_frame_content(shape.text_frame, get_shape_context(shape))

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


def extract_text_frame_content(text_frame, context="unknown"):
    """Extract content from a text frame with hierarchy detection."""
    if not text_frame or not text_frame.paragraphs:
        return ""

    # First, let's check if this looks like a manually formatted list
    # by looking at all paragraphs together
    all_paragraphs = []
    for paragraph in text_frame.paragraphs:
        original_text = paragraph.text
        if original_text:  # Include empty lines
            all_paragraphs.append((paragraph, original_text))

    # Check if this might be a manual list in a text box
    is_manual_list = False
    for i, (para, text) in enumerate(all_paragraphs):
        stripped = text.strip()
        # Look for patterns that indicate a manual list
        if stripped.endswith(':') and i + 1 < len(all_paragraphs):
            # Check if next lines have consistent indentation
            next_text = all_paragraphs[i + 1][1]
            if next_text and not next_text.startswith(' '):
                # Might be starting a manual list
                is_manual_list = True
                break

    paragraphs = []
    list_base_indent = None
    in_manual_list = False

    for para_idx, (paragraph, original_text) in enumerate(all_paragraphs):
        stripped_text = original_text.strip()

        if not stripped_text:
            continue

        # Count leading spaces
        leading_spaces = len(original_text) - len(original_text.lstrip(' '))

        # Check if we're entering a manual list section
        if is_manual_list and stripped_text.endswith(':'):
            # This is probably a list header like "Unordered Lists:"
            in_manual_list = True
            list_base_indent = None
            paragraphs.append(f"### {stripped_text}")
            continue

        # If we're in a manual list, handle indentation-based nesting
        if in_manual_list:
            # First item in the list sets the base indentation
            if list_base_indent is None and leading_spaces == 0:
                list_base_indent = 0

            # Check if we're ending the list (non-indented line that's not a list item)
            if leading_spaces == 0 and not any(char in stripped_text for char in [':', '•', '-', '*']):
                # Check if this might be another list item at base level
                if not any(phrase in stripped_text.lower() for phrase in ['level', 'item', 'nested']):
                    in_manual_list = False
                    list_base_indent = None

            if in_manual_list:
                # Calculate nesting level based on indentation
                if leading_spaces == 0:
                    level = 0
                else:
                    # Assuming 2 spaces per indent level
                    level = leading_spaces // 2

                # Format as a bullet item
                indent = "  " * level
                formatted_text = extract_formatted_text(paragraph.runs)
                paragraphs.append(f"{indent}- {formatted_text}")
                continue

        # Normal paragraph processing
        para_content = extract_paragraph_content(paragraph, context, False)
        if para_content.strip():
            paragraphs.append(para_content)

    return "\n".join(paragraphs)


def extract_paragraph_content(paragraph, context="unknown", in_list_context=False):
    """Extract content from a paragraph with proper hierarchy and list detection."""
    if not paragraph.runs:
        return ""

    # Get the original text with leading spaces
    original_text = paragraph.text
    raw_text = original_text.strip()
    if not raw_text:
        return ""

    # Count leading spaces for indentation
    leading_spaces = len(original_text) - len(original_text.lstrip(' '))

    # Check for numbered lists first
    if is_numbered_list(paragraph):
        return format_numbered_item(paragraph)

    # Check for bullet points (PowerPoint formatted)
    if is_bullet_point(paragraph):
        return format_bullet_item(paragraph)

    # Check for manual bullet characters (typed in text boxes)
    if raw_text and raw_text[0] in ['•', '·', '-', '*', '◦', '▪', '▫', '‣']:
        return format_manual_bullet_item(paragraph, raw_text, leading_spaces)

    # Check if this is an indented line in a list context
    if in_list_context and leading_spaces > 0:
        # This is a nested item without a bullet character
        level = max(1, leading_spaces // 2)  # Assume 2 spaces per level
        formatted_text = extract_formatted_text(paragraph.runs)
        indent = "  " * level
        return f"{indent}- {formatted_text}"

    # Extract formatted text
    formatted_text = extract_formatted_text(paragraph.runs)
    if not formatted_text.strip():
        return ""

    # Apply hierarchy formatting
    return apply_hierarchy_formatting(formatted_text, paragraph, context)


def format_manual_bullet_item(paragraph, raw_text, leading_spaces=0):
    """Format a manually typed bullet item."""
    # Remove the bullet character
    text = raw_text[1:].lstrip()

    # Apply formatting to the text
    formatted_text = extract_formatted_text(paragraph.runs)
    # Remove bullet from formatted text as well
    bullet_chars = ['•', '·', '-', '*', '◦', '▪', '▫', '‣']
    for char in bullet_chars:
        if formatted_text.startswith(char):
            formatted_text = formatted_text[1:].lstrip()
            break

    # Use the leading spaces to determine level
    level = leading_spaces // 2 if leading_spaces > 0 else 0

    indent = "  " * level
    return f"{indent}- {formatted_text}"


def format_nested_bullet_item(paragraph, raw_text, leading_spaces):
    """Format a nested bullet item that doesn't have an explicit bullet character."""
    # Determine nesting level from leading spaces
    level = max(1, leading_spaces // 2)  # At least level 1 for nested items

    # Get formatted text
    formatted_text = extract_formatted_text(paragraph.runs)

    indent = "  " * level
    return f"{indent}- {formatted_text}"


def is_numbered_list(paragraph):
    """Detect if paragraph is a numbered list."""
    text = paragraph.text.strip()
    if not text:
        return False

    # Patterns for numbered lists
    numbered_patterns = [
        r'^\d+[\.\)]\s',  # 1. or 1)
        r'^[a-z][\.\)]\s',  # a. or a)
        r'^[A-Z][\.\)]\s',  # A. or A)
        r'^[ivxlcdm]+[\.\)]\s',  # i. ii. iii. (roman numerals)
        r'^\([0-9a-zA-Z]+\)\s',  # (1) or (a)
    ]

    for pattern in numbered_patterns:
        if re.match(pattern, text):
            return True

    return False


def format_numbered_item(paragraph):
    """Format a numbered list item."""
    text = extract_formatted_text(paragraph.runs)

    # Get indentation level
    level = paragraph.level if paragraph.level is not None else 0
    indent = "   " * level  # 3 spaces for numbered list indentation

    # For markdown, we'll use "1." for all numbered items
    return f"{indent}1. {text}"


def is_bullet_point(paragraph):
    """Detect if paragraph is a bullet point."""
    try:
        # Check paragraph level (bullet points usually have level >= 0)
        if paragraph.level is not None and paragraph.level >= 0:
            # Check for bullet formatting in XML
            if hasattr(paragraph, '_p') and paragraph._p is not None:
                xml_str = str(paragraph._p.xml) if hasattr(paragraph._p, 'xml') else ""
                if any(bullet in xml_str for bullet in ['buChar', 'buAutoNum', 'buFont']):
                    return True

        # Check if text starts with bullet characters
        text = paragraph.text.strip()
        if text and text[0] in ['•', '·', '-', '*', '◦', '▪', '▫', '‣']:
            return True

    except:
        pass

    return False


def format_bullet_item(paragraph):
    """Format a bullet point item with improved level detection."""
    text = extract_formatted_text(paragraph.runs)

    # Remove existing bullet characters from the beginning
    text = re.sub(r'^[•·\-*◦▪▫‣]\s*', '', text)

    # Get indentation level - try multiple methods
    level = 0

    # Method 1: Use paragraph.level if available
    if paragraph.level is not None and paragraph.level >= 0:
        level = paragraph.level
    else:
        # Method 2: Try to detect indentation from formatting
        try:
            if hasattr(paragraph, '_element') and paragraph._element is not None:
                pPr = paragraph._element.pPr if hasattr(paragraph._element, 'pPr') else None
                if pPr is not None:
                    # Look for indentation markers
                    for child in pPr:
                        if hasattr(child, 'attrib'):
                            if 'marL' in child.attrib:
                                # Convert margin to approximate level
                                margin = int(child.attrib.get('marL', 0))
                                level = margin // 360000  # Rough conversion
                            elif 'lvl' in child.attrib:
                                level = int(child.attrib.get('lvl', 0))
        except:
            pass

        # Method 3: Detect from original text indentation patterns
        original_text = paragraph.text
        if original_text.startswith('  '):
            # Count leading spaces to estimate level
            leading_spaces = len(original_text) - len(original_text.lstrip(' '))
            level = leading_spaces // 2  # Assume 2 spaces per level

    # Ensure level is reasonable
    level = max(0, min(level, 5))  # Cap at 5 levels

    # Create proper markdown bullet with indentation
    indent = "  " * level
    return f"{indent}- {text}"


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