# Replace your bullet detection functions with these improved versions:

def detect_bullet_level_direct(paragraph):
    """Direct bullet level detection that avoids false positives from content."""

    # Method 1: Check PowerPoint's native bullet formatting FIRST
    # This is the most reliable method - if PowerPoint says it's a bullet, it is
    try:
        if hasattr(paragraph, 'level') and paragraph.level is not None:
            if is_actually_bullet_formatted(paragraph):
                return paragraph.level
    except:
        pass

    # Method 2: Check XML for bullet formatting
    xml_level = get_bullet_level_from_xml(paragraph)
    if xml_level >= 0:
        return xml_level

    # Method 3: Manual bullet detection - but be much more careful
    raw_text = paragraph.text
    stripped_text = raw_text.strip()

    if stripped_text and is_likely_manual_bullet(raw_text, stripped_text):
        # Count leading spaces to determine level
        leading_spaces = len(raw_text) - len(raw_text.lstrip(' '))
        if leading_spaces == 0:
            return 0
        else:
            return min(leading_spaces // 2, 6)  # Allow up to 6 levels

    return -1  # Not a bullet


def is_likely_manual_bullet(raw_text, stripped_text):
    """
    Determine if this is likely a manually typed bullet point.
    This function is much more careful to avoid false positives.
    """
    if not stripped_text:
        return False

    first_char = stripped_text[0]

    # Common manual bullet characters
    definite_bullets = ['•', '◦', '▪', '▫', '‣', '·']

    # Characters that could be bullets but are often used in content
    maybe_bullets = ['-', '*', '+', '→', '►', '✓', '✗']

    # Characters that are almost never bullets when used in content
    content_chars = ['α', 'β', 'γ', 'δ', '∑', '∫', '√', '∞', '≠', '≤', '≥',
                     '$', '€', '£', '¥', '₹', '₽', '₿', '←', '↑', '↓', '➜',
                     '★', '☆', '♥', '♦', '♣', '♠', '<', '>', '&', '{', '}',
                     '[', ']', '|', '\\', '/', '@', '#', '%', '^', '~', '`',
                     '±', '×', '÷', '²', '³', '¼', '½', '¾', 'π', 'θ', 'φ']

    # If it starts with a content character, it's definitely not a bullet
    if first_char in content_chars:
        return False

    # If it starts with a definite bullet character, check the context
    if first_char in definite_bullets:
        return is_bullet_context(stripped_text)

    # If it starts with a maybe-bullet character, be very careful
    if first_char in maybe_bullets:
        return is_bullet_context(stripped_text) and looks_like_bullet_formatting(raw_text)

    return False


def is_bullet_context(text):
    """
    Check if the text looks like a bullet point rather than content with symbols.
    """
    # If the text contains a colon early on, it's probably a content header
    # like "Mathematical symbols: α β γ δ"
    if ':' in text[:20]:
        return False

    # If it contains multiple special characters in sequence, it's probably content
    # like "α β γ δ ∑ ∫ √ ∞"
    special_char_count = 0
    for char in text[:30]:  # Check first 30 characters
        if ord(char) > 127:  # Non-ASCII character
            special_char_count += 1

    if special_char_count > 3:  # Too many special characters for a bullet
        return False

    # If it starts with "Mathematical", "Currency", "Technical", etc., it's content
    content_prefixes = ['mathematical', 'currency', 'technical', 'international',
                        'symbols', 'arrows', 'emojis', 'code', 'punctuation']

    text_lower = text.lower()
    if any(text_lower.startswith(prefix) for prefix in content_prefixes):
        return False

    return True


def looks_like_bullet_formatting(raw_text):
    """
    Check if the text formatting looks like a bullet point.
    Real bullets usually have consistent indentation patterns.
    """
    # Check leading whitespace pattern
    leading_spaces = len(raw_text) - len(raw_text.lstrip(' '))

    # Real bullets often have consistent indentation (0, 2, 4, 6, etc.)
    if leading_spaces % 2 == 0:
        return True

    # Single character followed by space is more likely a bullet
    stripped = raw_text.strip()
    if len(stripped) > 1 and stripped[1] == ' ':
        return True

    return False


def is_actually_bullet_formatted(paragraph):
    """Enhanced check for PowerPoint bullet formatting."""
    try:
        # Check the XML for bullet formatting indicators
        if hasattr(paragraph, '_p') and paragraph._p is not None:
            xml_str = str(paragraph._p.xml)

            # Look for bullet-related XML elements
            bullet_indicators = [
                'buChar',  # Custom bullet character
                'buAutoNum',  # Auto-numbered bullets
                'buFont',  # Bullet font
                'buSzPct',  # Bullet size
                'buClr',  # Bullet color
            ]

            # Count how many bullet indicators we find
            indicator_count = sum(1 for indicator in bullet_indicators if indicator in xml_str)

            # If we find multiple indicators, it's definitely a bullet
            if indicator_count >= 2:
                return True

            # If we find at least one and there's a level indicator, probably a bullet
            if indicator_count >= 1 and 'lvl=' in xml_str:
                return True

    except:
        pass

    return False


def get_bullet_level_from_xml(paragraph):
    """Extract bullet level from XML with better validation."""
    try:
        if hasattr(paragraph, '_element') and paragraph._element is not None:
            pPr = getattr(paragraph._element, 'pPr', None)
            if pPr is not None:
                for child in pPr:
                    attribs = getattr(child, 'attrib', {})

                    # Look for explicit level attribute
                    if 'lvl' in attribs:
                        level = int(attribs['lvl'])
                        # Only return if we also have bullet formatting
                        if is_actually_bullet_formatted(paragraph):
                            return level

                    # Look for margin-based level
                    if 'marL' in attribs:
                        margin = int(attribs.get('marL', 0))
                        # Only consider it a bullet if margin is substantial and we have bullet formatting
                        if margin >= 360000 and is_actually_bullet_formatted(paragraph):  # At least 0.5 inch
                            level = margin // 720000  # Convert to level
                            return min(level, 6)

    except:
        pass

    return -1


# Update the debug function to show more details
def debug_bullet_detection_simple(paragraph):
    """Enhanced debug function."""
    raw_text = paragraph.text
    stripped_text = raw_text.strip()

    print(f"Text: '{stripped_text[:50]}{'...' if len(stripped_text) > 50 else ''}'")
    print(f"First char: '{stripped_text[0] if stripped_text else 'EMPTY'}'")
    print(f"PowerPoint level: {getattr(paragraph, 'level', 'None')}")
    print(f"Leading spaces: {len(raw_text) - len(raw_text.lstrip(' '))}")
    print(f"Is bullet formatted: {is_actually_bullet_formatted(paragraph)}")
    print(f"XML level: {get_bullet_level_from_xml(paragraph)}")

    if stripped_text:
        print(f"Is likely manual bullet: {is_likely_manual_bullet(raw_text, stripped_text)}")
        print(f"Is bullet context: {is_bullet_context(stripped_text)}")
        print(f"Looks like bullet formatting: {looks_like_bullet_formatting(raw_text)}")

    print(f"Final detected level: {detect_bullet_level_direct(paragraph)}")
    print("---")