"""
Markdown Converter - Converts structured PowerPoint data to markdown format
Handles text formatting, tables, images, charts, and slide title conversion
"""

import re


class MarkdownConverter:
    """
    Converts structured presentation data to clean markdown format.
    Handles proper formatting, bullets, tables, and slide title detection.
    """

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """
        Convert structured presentation data to markdown.

        Args:
            data (dict): Structured presentation data
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1

        Returns:
            str: Formatted markdown content
        """
        markdown_parts = []

        for slide in data["slides"]:
            # Add slide marker
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            for block in slide["content_blocks"]:
                if block["type"] == "text":
                    markdown_parts.append(self._convert_text_block_to_markdown(block))
                elif block["type"] == "table":
                    markdown_parts.append(self._convert_table_to_markdown(block))
                elif block["type"] == "image":
                    markdown_parts.append(self._convert_image_to_markdown(block))
                elif block["type"] == "chart":
                    markdown_parts.append(self._convert_chart_to_markdown(block))
                elif block["type"] == "group":
                    markdown_parts.append(self._convert_group_to_markdown(block))

        markdown_content = "\n\n".join(filter(None, markdown_parts))

        # Post-process to convert slide titles if requested
        if convert_slide_titles:
            markdown_content = self._convert_slide_titles_to_headings(markdown_content)

        return markdown_content

    def _convert_text_block_to_markdown(self, block):
        """
        Convert text block to markdown with proper formatting.

        Args:
            block (dict): Text content block

        Returns:
            str: Formatted markdown text
        """
        lines = []

        for para in block["paragraphs"]:
            line = self._convert_paragraph_to_markdown(para)
            if line:
                lines.append(line)

        # Apply shape-level hyperlink if present
        result = "\n".join(lines)
        if block.get("shape_hyperlink") and result:
            result = f"[{result}]({block['shape_hyperlink']})"

        return result

    def _convert_paragraph_to_markdown(self, para):
        """
        Convert paragraph to markdown with correct formatting.

        Args:
            para (dict): Paragraph data with formatting

        Returns:
            str: Formatted markdown paragraph
        """
        if not para.get("clean_text"):
            return ""

        # Build formatted text from runs
        formatted_text = self._build_formatted_text_from_runs(
            para["formatted_runs"], para["clean_text"]
        )

        # Apply structural formatting based on hints
        hints = para.get("hints", {})

        # Handle bullets
        if hints.get("is_bullet", False):
            level = hints.get("bullet_level", 0)
            if level < 0:
                level = 0
            indent = "  " * level
            return f"{indent}- {formatted_text}"

        # Handle numbered lists
        elif hints.get("is_numbered", False):
            return f"1. {formatted_text}"

        # Handle headings
        elif hints.get("likely_heading", False):
            # Determine heading level
            if hints.get("all_caps") or len(para["clean_text"]) < 30:
                return f"## {formatted_text}"
            else:
                return f"### {formatted_text}"

        # Regular paragraph
        else:
            return formatted_text

    def _build_formatted_text_from_runs(self, runs, clean_text):
        """
        Build formatted text from runs, handling text formatting.

        Args:
            runs (list): List of formatted text runs
            clean_text (str): Clean text content

        Returns:
            str: Formatted markdown text
        """
        if not runs:
            return clean_text

        # Check if we have any formatting
        has_formatting = any(
            run.get("bold") or run.get("italic") or run.get("hyperlink")
            for run in runs
        )

        if not has_formatting:
            return clean_text

        # Build formatted text preserving run boundaries
        formatted_parts = []

        for run in runs:
            text = run["text"]
            if not text:
                continue

            # Apply text formatting
            if run.get("bold") and run.get("italic"):
                text = f"***{text}***"
            elif run.get("bold"):
                text = f"**{text}**"
            elif run.get("italic"):
                text = f"*{text}*"

            # Apply hyperlink
            if run.get("hyperlink"):
                text = f"[{text}]({run['hyperlink']})"

            formatted_parts.append(text)

        return "".join(formatted_parts)

    def _convert_table_to_markdown(self, block):
        """
        Convert table to markdown format.

        Args:
            block (dict): Table content block

        Returns:
            str: Markdown table
        """
        if not block["data"]:
            return ""

        markdown = ""
        for i, row in enumerate(block["data"]):
            # Escape pipes in cell content
            escaped_row = [cell.replace("|", "\\|") for cell in row]
            markdown += "| " + " | ".join(escaped_row) + " |\n"

            # Add separator after header row
            if i == 0:
                markdown += "| " + " | ".join("---" for _ in row) + " |\n"

        return markdown

    def _convert_image_to_markdown(self, block):
        """
        Convert image to markdown format.

        Args:
            block (dict): Image content block

        Returns:
            str: Markdown image
        """
        image_md = f"![{block['alt_text']}](image)"

        if block.get("hyperlink"):
            image_md = f"[{image_md}]({block['hyperlink']})"

        return image_md

    def _convert_chart_to_markdown(self, block):
        """
        Convert chart to markdown with diagram candidate comment.

        Args:
            block (dict): Chart content block

        Returns:
            str: Markdown chart representation
        """
        chart_md = f"**Chart: {block.get('title', 'Untitled Chart')}**\n"
        chart_md += f"*Chart Type: {block.get('chart_type', 'unknown')}*\n\n"

        # Add data if available
        if block.get('categories') and block.get('series'):
            chart_md += "Data:\n"
            for series in block['series']:
                if series.get('name'):
                    chart_md += f"- {series['name']}: "
                    if series.get('values'):
                        chart_md += ", ".join(map(str, series['values'][:5]))
                        if len(series['values']) > 5:
                            chart_md += "..."
                    chart_md += "\n"

        # Add diagram candidate comment for Claude processing
        chart_md += f"\n<!-- DIAGRAM_CANDIDATE: chart, type={block.get('chart_type', 'unknown')} -->\n"

        if block.get("hyperlink"):
            chart_md = f"[{chart_md}]({block['hyperlink']})"

        return chart_md

    def _convert_group_to_markdown(self, block):
        """
        Convert grouped shapes to markdown.

        Args:
            block (dict): Group content block

        Returns:
            str: Markdown representation of group
        """
        extracted_blocks = block.get("extracted_blocks", [])

        if not extracted_blocks:
            return ""

        # Convert each extracted block to markdown
        content_parts = []

        for extracted_block in extracted_blocks:
            if extracted_block["type"] == "text":
                content = self._convert_text_block_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "image":
                content = self._convert_image_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "table":
                content = self._convert_table_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "chart":
                content = self._convert_chart_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "line":
                # Lines don't produce visible content but are tracked for diagram analysis
                pass
            elif extracted_block["type"] == "arrow":
                # Arrows don't produce visible content but are tracked for diagram analysis
                pass
            elif extracted_block["type"] == "shape":
                # Generic shapes might have minimal content
                content = f"[Shape: {extracted_block.get('shape_subtype', 'unknown')}]"
                content_parts.append(content)

        # Join all content
        group_md = "\n\n".join(content_parts) if content_parts else ""

        # Add shape-level hyperlink if present
        if block.get("hyperlink") and group_md:
            group_md = f"[{group_md}]({block['hyperlink']})"

        return group_md

    def _convert_slide_titles_to_headings(self, markdown_content):
        """
        Post-process markdown to convert slide titles from bullet points to H1 headings.

        Args:
            markdown_content (str): Original markdown content

        Returns:
            str: Markdown with converted slide titles
        """
        lines = markdown_content.split('\n')
        processed_lines = []

        i = 0
        while i < len(lines):
            line = lines[i]
            processed_lines.append(line)

            # Check if this is a slide marker
            if line.strip().startswith('<!-- Slide ') and line.strip().endswith(' -->'):
                # Look ahead for the first non-empty content line
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    processed_lines.append(lines[j])
                    j += 1

                # Check if the next content line is a bullet that looks like a title
                if j < len(lines):
                    next_line = lines[j].strip()
                    if self._is_likely_slide_title(next_line):
                        # Convert bullet to H1 heading
                        title_text = self._extract_title_from_bullet(next_line)
                        processed_lines.append(f"\n# {title_text}")
                        i = j  # Skip the original bullet line
                    else:
                        i = j - 1  # Process the next line normally
                else:
                    break

            i += 1

        return '\n'.join(processed_lines)

    def _is_likely_slide_title(self, line):
        """
        Determine if a line is likely a slide title.

        Args:
            line (str): Line to evaluate

        Returns:
            bool: True if the line appears to be a slide title
        """
        if not line.strip():
            return False

        # Must be a bullet point to be converted
        if not line.startswith('- '):
            return False

        # Extract the text content
        text_content = line[2:].strip()

        # Title characteristics
        title_indicators = [
            len(text_content) <= 150,  # Reasonable title length
            not text_content.endswith(('.', '!', '?', ';', ':')),  # No ending punctuation
            not self._contains_multiple_sentences(text_content),  # Single phrases
            not text_content.lower().startswith(('the following', 'here are', 'this slide', 'key points')),
        ]

        # Additional positive indicators
        positive_indicators = [
            text_content.isupper(),  # All caps suggests title
            text_content.istitle(),  # Title case suggests title
            len(text_content.split()) <= 10,  # Short phrases are more likely titles
            any(word in text_content.lower() for word in
                ['overview', 'introduction', 'conclusion', 'agenda', 'objectives']),
        ]

        # Must meet basic criteria and have positive indicator
        basic_criteria_met = all(title_indicators)
        has_positive_indicator = any(positive_indicators)

        return basic_criteria_met and (has_positive_indicator or len(text_content.split()) <= 6)

    def _extract_title_from_bullet(self, bullet_line):
        """
        Extract clean title text from a bullet point line.

        Args:
            bullet_line (str): Bullet point line

        Returns:
            str: Clean title text
        """
        # Remove bullet prefix
        title_text = bullet_line[2:].strip()

        # Clean up common title artifacts
        title_text = title_text.strip('*_`')  # Remove markdown formatting artifacts

        return title_text

    def _contains_multiple_sentences(self, text):
        """
        Check if text contains multiple sentences.

        Args:
            text (str): Text to check

        Returns:
            bool: True if text appears to contain multiple sentences
        """
        # Simple heuristic: look for sentence-ending punctuation followed by space and capital letter
        sentence_pattern = r'[.!?]\s+[A-Z]'
        return bool(re.search(sentence_pattern, text))

