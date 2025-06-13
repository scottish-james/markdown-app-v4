"""
Text Processor - Handles advanced text formatting and bullet detection
Specializes in extracting formatted text with proper bullet hierarchies
"""

import re


class TextProcessor:
    """
    Processes text content from PowerPoint shapes with advanced formatting detection.
    Handles bullets, numbering, hyperlinks, and text formatting (bold, italic).
    """

    def extract_text_frame(self, text_frame, shape):
        """
        Extract text content from a text frame with proper formatting.

        Args:
            text_frame: python-pptx TextFrame object
            shape: parent Shape object

        Returns:
            dict: Text content block with formatting
        """
        if not text_frame.paragraphs:
            return None

        block = {
            "type": "text",
            "paragraphs": [],
            "shape_hyperlink": self._extract_shape_hyperlink(shape)
        }

        for para_idx, para in enumerate(text_frame.paragraphs):
            if not para.text.strip():
                continue

            para_data = self.process_paragraph(para)
            if para_data:
                block["paragraphs"].append(para_data)

        return block if block["paragraphs"] else None

    def extract_plain_text(self, shape):
        """
        Extract plain text from shape with basic analysis.

        Args:
            shape: python-pptx Shape object

        Returns:
            dict: Text content block
        """
        if not hasattr(shape, 'text') or not shape.text:
            return None

        return {
            "type": "text",
            "paragraphs": [{
                "raw_text": shape.text,
                "clean_text": shape.text.strip(),
                "formatted_runs": [{"text": shape.text, "bold": False, "italic": False, "hyperlink": None}],
                "hints": self._analyze_plain_text_hints(shape.text)
            }],
            "shape_hyperlink": self._extract_shape_hyperlink(shape)
        }

    def process_paragraph(self, para):
        """
        Process a single paragraph with advanced bullet detection and formatting.

        Args:
            para: python-pptx Paragraph object

        Returns:
            dict: Processed paragraph data
        """
        raw_text = para.text
        if not raw_text.strip():
            return None

        # Check PowerPoint's bullet formatting
        ppt_level = getattr(para, 'level', None)
        is_ppt_bullet, xml_level = self._check_xml_bullet_formatting(para)

        # Determine final bullet level
        bullet_level = self._determine_bullet_level(is_ppt_bullet, xml_level, ppt_level)

        # Check for manual bullets and numbering
        clean_text = raw_text.strip()
        manual_bullet = self._is_manual_bullet(clean_text)
        numbered = self._is_numbered_list(clean_text)

        # Process text based on formatting
        if manual_bullet and bullet_level < 0:
            # Estimate level from indentation
            leading_spaces = len(raw_text) - len(raw_text.lstrip())
            bullet_level = min(leading_spaces // 2, 6)
            clean_text = self._remove_bullet_char(clean_text)
        elif bullet_level >= 0:
            # Remove manual bullet chars if PowerPoint formatted it
            clean_text = self._remove_bullet_char(clean_text)
        elif numbered:
            clean_text = self._remove_number_prefix(clean_text)

        # Extract formatted runs
        formatted_runs = self._extract_runs_with_text_preservation(
            para.runs, clean_text, bullet_level >= 0 or numbered
        )

        para_data = {
            "raw_text": raw_text,
            "clean_text": clean_text,
            "formatted_runs": formatted_runs,
            "hints": {
                "has_powerpoint_level": ppt_level is not None,
                "powerpoint_level": ppt_level,
                "bullet_level": bullet_level,
                "is_bullet": bullet_level >= 0,
                "is_numbered": numbered,
                "starts_with_bullet": manual_bullet,
                "starts_with_number": numbered,
                "short_text": len(clean_text) < 100,
                "all_caps": clean_text.isupper() if clean_text else False,
                "likely_heading": self._is_likely_heading(clean_text)
            }
        }

        return para_data

    def _check_xml_bullet_formatting(self, para):
        """
        Check XML for bullet formatting indicators.

        Args:
            para: python-pptx Paragraph object

        Returns:
            tuple: (is_bullet, xml_level)
        """
        is_ppt_bullet = False
        xml_level = None

        try:
            if hasattr(para, '_p') and para._p is not None:
                xml_str = str(para._p.xml)
                # Look for bullet indicators
                if any(indicator in xml_str for indicator in ['buChar', 'buAutoNum', 'buFont']):
                    is_ppt_bullet = True
                    # Try to extract level
                    level_match = re.search(r'lvl="(\d+)"', xml_str)
                    if level_match:
                        xml_level = int(level_match.group(1))
        except:
            pass

        return is_ppt_bullet, xml_level

    def _determine_bullet_level(self, is_ppt_bullet, xml_level, ppt_level):
        """
        Determine the final bullet level from various sources.

        Args:
            is_ppt_bullet (bool): Whether XML indicates bullet
            xml_level (int): Level from XML
            ppt_level (int): Level from PowerPoint

        Returns:
            int: Final bullet level (-1 if not a bullet)
        """
        bullet_level = -1

        if is_ppt_bullet:
            bullet_level = xml_level if xml_level is not None else (ppt_level if ppt_level is not None else 0)
        elif ppt_level is not None:
            # PowerPoint says it has a level, trust it
            bullet_level = ppt_level

        return bullet_level

    def _extract_runs_with_text_preservation(self, runs, clean_text, has_prefix_removed):
        """
        Extract formatted runs while preserving formatting after bullet/number removal.

        Args:
            runs: List of python-pptx Run objects
            clean_text (str): Text with prefixes removed
            has_prefix_removed (bool): Whether a prefix was removed

        Returns:
            list: Formatted run data
        """
        if not runs:
            return [{"text": clean_text, "bold": False, "italic": False, "hyperlink": None}]

        formatted_runs = []

        if has_prefix_removed:
            # Find where the clean text starts in the original runs
            full_text = "".join(run.text for run in runs)
            start_pos = self._find_clean_text_start_position(full_text, clean_text)

            # Process runs, skipping content before start_pos
            char_count = 0
            for run in runs:
                run_text = run.text
                run_start = char_count
                run_end = char_count + len(run_text)

                # Skip if this run is entirely before our clean text
                if run_end <= start_pos:
                    char_count += len(run_text)
                    continue

                # Adjust text if run spans the start position
                if run_start < start_pos < run_end:
                    run_text = run_text[start_pos - run_start:]

                if run_text:
                    formatted_runs.append(self._extract_run_formatting(run, run_text))

                char_count += len(run.text)
        else:
            # No prefix removed, process runs normally
            for run in runs:
                if run.text:
                    formatted_runs.append(self._extract_run_formatting(run, run.text))

        return formatted_runs

    def _find_clean_text_start_position(self, full_text, clean_text):
        """
        Find where clean text starts in the original full text.

        Args:
            full_text (str): Original text with prefixes
            clean_text (str): Text with prefixes removed

        Returns:
            int: Start position of clean text
        """
        # Try to find clean text position
        for i in range(len(full_text)):
            remaining = full_text[i:].strip()
            if remaining == clean_text:
                return i
        return 0  # Fallback

    def _extract_run_formatting(self, run, text_override=None):
        """
        Extract formatting from a single text run.

        Args:
            run: python-pptx Run object
            text_override (str): Override text content

        Returns:
            dict: Run formatting data
        """
        run_data = {
            "text": text_override if text_override is not None else run.text,
            "bold": False,
            "italic": False,
            "hyperlink": None
        }

        # Get formatting
        try:
            font = run.font
            if hasattr(font, 'bold') and font.bold:
                run_data["bold"] = True
            if hasattr(font, 'italic') and font.italic:
                run_data["italic"] = True
        except:
            pass

        # Get hyperlinks
        try:
            if hasattr(run, 'hyperlink') and run.hyperlink and run.hyperlink.address:
                run_data["hyperlink"] = self._fix_url(run.hyperlink.address)
        except:
            pass

        return run_data

    def _is_manual_bullet(self, text):
        """
        Check if text starts with a manual bullet character.

        Args:
            text (str): Text to check

        Returns:
            bool: True if text starts with bullet character
        """
        if not text:
            return False
        bullet_chars = '•◦▪▫‣·○■□→►✓✗-*+※◆◇'
        return text[0] in bullet_chars

    def _is_numbered_list(self, text):
        """
        Check if text starts with a number pattern.

        Args:
            text (str): Text to check

        Returns:
            bool: True if text starts with numbering
        """
        patterns = [
            r'^\d+[\.\)]\s+',  # 1. or 1)
            r'^[a-zA-Z][\.\)]\s+',  # a. or A)
            r'^[ivxlcdm]+[\.\)]\s+',  # Roman numerals (lowercase)
            r'^[IVXLCDM]+[\.\)]\s+',  # Roman numerals (uppercase)
        ]
        return any(re.match(pattern, text) for pattern in patterns)

    def _remove_bullet_char(self, text):
        """
        Remove bullet characters from start of text.

        Args:
            text (str): Text with bullet characters

        Returns:
            str: Text without bullet characters
        """
        if not text:
            return text
        return re.sub(r'^[•◦▪▫‣·○■□→►✓✗\-\*\+※◆◇]\s*', '', text)

    def _remove_number_prefix(self, text):
        """
        Remove number prefix from text.

        Args:
            text (str): Text with number prefix

        Returns:
            str: Text without number prefix
        """
        return re.sub(r'^[^\s]+\s+', '', text)

    def _is_likely_heading(self, text):
        """
        Determine if text is likely a heading.

        Args:
            text (str): Text to analyze

        Returns:
            bool: True if text appears to be a heading
        """
        if not text or len(text) > 150:
            return False

        # All caps
        if text.isupper() and len(text) > 2:
            return True

        # Short text without ending punctuation
        if len(text) < 80 and not text.endswith(('.', '!', '?', ';', ':', ',')):
            return True

        return False

    def _analyze_plain_text_hints(self, text):
        """
        Analyze plain text for formatting hints.

        Args:
            text (str): Text to analyze

        Returns:
            dict: Analysis hints
        """
        if not text:
            return {}

        stripped = text.strip()

        # Check each line for bullets
        lines = text.split('\n')
        has_bullets = any(line.strip() and self._is_manual_bullet(line.strip()) for line in lines)

        return {
            "has_powerpoint_level": False,
            "powerpoint_level": None,
            "bullet_level": -1,
            "is_bullet": has_bullets,
            "is_numbered": any(self._is_numbered_list(line.strip()) for line in lines if line.strip()),
            "starts_with_bullet": stripped and self._is_manual_bullet(stripped),
            "starts_with_number": bool(re.match(r'^\s*\d+[\.\)]\s', text)),
            "short_text": len(stripped) < 100,
            "all_caps": stripped.isupper() if stripped else False,
            "likely_heading": self._is_likely_heading(stripped)
        }

    def _extract_shape_hyperlink(self, shape):
        """
        Extract shape-level hyperlink if present.

        Args:
            shape: python-pptx Shape object

        Returns:
            str: URL or None if no hyperlink
        """
        try:
            if hasattr(shape, 'click_action') and shape.click_action:
                if hasattr(shape.click_action, 'hyperlink') and shape.click_action.hyperlink:
                    if shape.click_action.hyperlink.address:
                        return self._fix_url(shape.click_action.hyperlink.address)
        except:
            pass
        return None

    def _fix_url(self, url):
        """
        Fix URLs by adding schemes if missing.

        Args:
            url (str): URL to fix

        Returns:
            str: Properly formatted URL
        """
        if not url:
            return url

        if '@' in url and not url.startswith('mailto:'):
            return f"mailto:{url}"

        if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
            if url.startswith('www.') or any(
                    domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                return f"https://{url}"

        return url