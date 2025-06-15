"""
Text Processor - Handles advanced text formatting with XML-driven detection
Simplified to rely on XML data instead of text pattern matching
Fixed: Bold detection for consistently formatted text
"""

import re


class TextProcessor:
    """
    Processes text content from PowerPoint shapes using XML data when available.
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
        Process a single paragraph using XML data for formatting detection.

        Args:
            para: python-pptx Paragraph object

        Returns:
            dict: Processed paragraph data
        """
        raw_text = para.text
        if not raw_text.strip():
            return None

        # Get formatting from XML (reliable source)
        ppt_level = getattr(para, 'level', None)
        is_ppt_bullet, xml_level = self._check_xml_bullet_formatting(para)

        # Determine final bullet level from XML data
        bullet_level = self._determine_bullet_level(is_ppt_bullet, xml_level, ppt_level)

        # Clean text based on XML formatting info
        clean_text = raw_text.strip()
        if bullet_level >= 0:
            # XML says it's a bullet/number, clean it up
            clean_text = self._remove_bullet_char(clean_text)

        # Extract formatted runs
        formatted_runs = self._extract_runs_with_formatting(para.runs, clean_text, bullet_level >= 0)

        # Detect headings from font size rather than text patterns
        likely_heading = self._is_likely_heading_from_font_size(para)

        para_data = {
            "raw_text": raw_text,
            "clean_text": clean_text,
            "formatted_runs": formatted_runs,
            "hints": {
                "has_powerpoint_level": ppt_level is not None,
                "powerpoint_level": ppt_level,
                "bullet_level": bullet_level,
                "is_bullet": bullet_level >= 0,
                "is_numbered": self._is_numbered_from_xml(para),
                "short_text": len(clean_text) < 100,
                "all_caps": clean_text.isupper() if clean_text else False,
                "likely_heading": likely_heading
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

    def _is_numbered_from_xml(self, para):
        """
        Check if paragraph is numbered based on XML data.

        Args:
            para: python-pptx Paragraph object

        Returns:
            bool: True if numbered list
        """
        try:
            if hasattr(para, '_p') and para._p is not None:
                xml_str = str(para._p.xml)
                # Look for numbering indicators in XML
                return 'buAutoNum' in xml_str
        except:
            return False

    def _is_likely_heading_from_font_size(self, para):
        """
        Determine if text is likely a heading based on font size and length from XML.

        Args:
            para: python-pptx Paragraph object

        Returns:
            bool: True if text appears to be a heading
        """
        try:
            text = para.text.strip()
            if not text or len(text) > 150:  # Keep length check - very long text unlikely to be heading
                return False

            # Get font size from first run
            if para.runs:
                first_run = para.runs[0]
                if hasattr(first_run, 'font') and hasattr(first_run.font, 'size') and first_run.font.size:
                    font_size_pt = first_run.font.size.pt
                    # Heading if font is 14pt or larger (typical heading threshold)
                    return font_size_pt >= 14

            # Fallback: if no font size available, use length only
            return len(text) < 80

        except Exception:
            # If we can't get font info, use conservative length check
            text = para.text.strip() if hasattr(para, 'text') else ""
            return len(text) < 80 and len(text) > 0

    def _determine_bullet_level(self, is_ppt_bullet, xml_level, ppt_level):
        """
        Determine the final bullet level from XML sources.

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

    def _extract_runs_with_formatting(self, runs, clean_text, has_prefix_removed):
        """
        Extract formatted runs while preserving formatting.

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

    def _analyze_plain_text_hints(self, text):
        """
        Analyze plain text for basic formatting hints (fallback for non-XML cases).

        Args:
            text (str): Text to analyze

        Returns:
            dict: Analysis hints
        """
        if not text:
            return {}

        stripped = text.strip()

        return {
            "has_powerpoint_level": False,
            "powerpoint_level": None,
            "bullet_level": -1,
            "is_bullet": False,
            "is_numbered": False,
            "starts_with_bullet": False,
            "starts_with_number": False,
            "short_text": len(stripped) < 100,
            "all_caps": stripped.isupper() if stripped else False,
            "likely_heading": len(stripped) < 80 and len(stripped) > 0
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


