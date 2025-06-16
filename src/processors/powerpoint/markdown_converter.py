"""
Markdown Converter - Updated to use XML-based semantic roles for title detection
No more pattern matching - trusts the XML semantic analysis completely

ARCHITECTURE OVERVIEW:
This component converts the structured data from ContentExtractor into clean,
readable markdown format. It now trusts the semantic role information from
XML analysis rather than trying to guess titles from text patterns.

TITLE DETECTION STRATEGY:
- XML semantic role "title" → H1 heading (no bullet point)
- XML semantic role "subtitle" → H2 heading (no bullet point)
- All other content → processed normally with bullets/formatting

SEMANTIC TRUST:
The AccessibilityOrderExtractor already did the hard work of XML analysis
to identify titles. We trust that completely and don't second-guess it
with text pattern matching.
"""

import re


class MarkdownConverter:
    """
    Converts structured presentation data to clean markdown format using XML semantic roles.
    """

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """
        Main conversion entry point - processes entire presentation.

        SEMANTIC-BASED PROCESSING:
        - Titles are already identified by XML analysis in semantic_role field
        - No more text pattern guessing - trust the XML analysis
        - Titles become H1 headings directly, not bullet points

        Args:
            data (dict): Structured presentation data from ContentExtractor
            convert_slide_titles (bool): Keep for compatibility, but XML controls titles now

        Returns:
            str: Complete markdown document with semantic-based title detection
        """
        markdown_parts = []

        for slide in data["slides"]:
            # Add slide marker for debugging
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            # Process each content block with semantic role awareness
            for block in slide["content_blocks"]:
                if block["type"] == "text":
                    # NEW: Check semantic role for direct title conversion
                    markdown_parts.append(self._convert_text_block_to_markdown(block))
                elif block["type"] == "table":
                    markdown_parts.append(self._convert_table_to_markdown(block))
                elif block["type"] == "image":
                    markdown_parts.append(self._convert_image_to_markdown(block))
                elif block["type"] == "chart":
                    markdown_parts.append(self._convert_chart_to_markdown(block))
                elif block["type"] == "group":
                    markdown_parts.append(self._convert_group_to_markdown(block))

        # Combine all parts with proper spacing
        markdown_content = "\n\n".join(filter(None, markdown_parts))

        # NEW: No post-processing needed - titles already converted based on semantic roles
        return markdown_content

    def _convert_text_block_to_markdown(self, block):
        """
        Convert text content blocks to markdown using semantic role information.

        SEMANTIC PROCESSING:
        - semantic_role "title" → H1 heading
        - semantic_role "subtitle" → H2 heading
        - All other content → normal paragraph processing

        Args:
            block (dict): Text content block with semantic role and paragraphs

        Returns:
            str: Formatted markdown text with proper semantic structure
        """
        lines = []
        semantic_role = block.get("semantic_role", "other")

        # NEW: Handle semantic roles directly without pattern matching
        if semantic_role == "title":
            # Titles become H1 headings directly
            for para in block["paragraphs"]:
                if para.get("clean_text"):
                    formatted_text = self._build_formatted_text_from_runs(
                        para["formatted_runs"], para["clean_text"]
                    )
                    lines.append(f"# {formatted_text}")
        elif semantic_role == "subtitle":
            # Subtitles become H2 headings directly
            for para in block["paragraphs"]:
                if para.get("clean_text"):
                    formatted_text = self._build_formatted_text_from_runs(
                        para["formatted_runs"], para["clean_text"]
                    )
                    lines.append(f"## {formatted_text}")
        else:
            # Process all other content normally
            for para in block["paragraphs"]:
                line = self._convert_paragraph_to_markdown(para)
                if line:
                    lines.append(line)

        # Combine paragraphs with newline separation
        result = "\n".join(lines)

        # Apply shape-level hyperlink if present
        if block.get("shape_hyperlink") and result:
            result = f"[{result}]({block['shape_hyperlink']})"

        return result

    def _convert_paragraph_to_markdown(self, para):
        """
        Convert individual paragraphs to markdown with structure and formatting.

        SIMPLIFIED PROCESSING:
        Since titles are handled at the block level using semantic roles,
        this method only deals with content paragraphs (bullets, text, etc.)
        """
        if not para.get("clean_text"):
            return ""

        # Build formatted text from individual runs
        formatted_text = self._build_formatted_text_from_runs(
            para["formatted_runs"], para["clean_text"]
        )

        # Apply structural formatting based on paragraph hints
        hints = para.get("hints", {})

        # Handle bullet points with proper indentation
        if hints.get("is_bullet", False):
            level = hints.get("bullet_level", 0)
            if level < 0:
                level = 0  # Ensure non-negative indentation
            indent = "  " * level  # 2 spaces per indentation level
            return f"{indent}- {formatted_text}"

        # Handle numbered lists
        elif hints.get("is_numbered", False):
            return f"1. {formatted_text}"

        # Handle remaining headings (should be rare since titles are handled semantically)
        elif hints.get("likely_heading", False):
            if hints.get("all_caps") or len(para["clean_text"]) < 30:
                return f"## {formatted_text}"  # H2 for short/caps headings
            else:
                return f"### {formatted_text}"  # H3 for longer headings

        # Regular paragraph
        else:
            return formatted_text

    def _convert_group_to_markdown(self, block):
        """
        Convert grouped shapes to markdown by processing extracted content with semantic awareness.
        """
        extracted_blocks = block.get("extracted_blocks", [])

        if not extracted_blocks:
            return ""

        # Process each extracted block with semantic role handling
        content_parts = []

        for extracted_block in extracted_blocks:
            content = None

            if extracted_block["type"] == "text":
                # Handle semantic roles in group content too
                content = self._convert_text_block_to_markdown(extracted_block)
            elif extracted_block["type"] == "image":
                content = self._convert_image_to_markdown(extracted_block)
            elif extracted_block["type"] == "table":
                content = self._convert_table_to_markdown(extracted_block)
            elif extracted_block["type"] == "chart":
                content = self._convert_chart_to_markdown(extracted_block)
            elif extracted_block["type"] == "line":
                # Lines are tracked for diagram analysis but don't produce visible content
                pass
            elif extracted_block["type"] == "arrow":
                # Arrows are tracked for diagram analysis but don't produce visible content
                pass
            elif extracted_block["type"] == "shape":
                # Generic shapes get basic representation
                content = f"[Shape: {extracted_block.get('shape_subtype', 'unknown')}]"

            if content:
                content_parts.append(content)

        # Combine all content with proper spacing
        group_md = "\n\n".join(content_parts) if content_parts else ""

        # Apply shape-level hyperlink if present
        if block.get("hyperlink") and group_md:
            group_md = f"[{group_md}]({block['hyperlink']})"

        return group_md

    def _build_formatted_text_from_runs(self, runs, clean_text):
        """
        Build formatted text from runs with FIXED consistent formatting handling.
        """
        if not runs:
            return clean_text

        # Filter out empty runs for formatting analysis
        text_runs = [run for run in runs if run.get("text")]

        if not text_runs:
            return clean_text

        # Analyze formatting consistency across all runs
        all_bold = all(run.get("bold", False) for run in text_runs)
        all_italic = all(run.get("italic", False) for run in text_runs)
        all_have_hyperlinks = all(run.get("hyperlink") for run in text_runs)

        # Check if all hyperlinks are the same
        if all_have_hyperlinks:
            unique_hyperlinks = set(run.get("hyperlink") for run in text_runs)
            all_same_hyperlink = len(unique_hyperlinks) == 1
        else:
            all_same_hyperlink = False

        # Apply consistent formatting to entire text (FIXED approach)
        if all_bold and all_italic and not all_same_hyperlink:
            return f"***{clean_text}***"  # Bold + italic
        elif all_bold and not all_same_hyperlink:
            return f"**{clean_text}**"  # Fixed: Single application of bold
        elif all_italic and not all_same_hyperlink:
            return f"*{clean_text}*"  # Single application of italic
        elif all_same_hyperlink:
            # All runs have same hyperlink - apply to entire text
            hyperlink = text_runs[0]["hyperlink"]
            if all_bold and all_italic:
                return f"[***{clean_text}***]({hyperlink})"
            elif all_bold:
                return f"[**{clean_text}**]({hyperlink})"
            elif all_italic:
                return f"[*{clean_text}*]({hyperlink})"
            else:
                return f"[{clean_text}]({hyperlink})"

        # Mixed formatting - use per-run logic (original approach)
        formatted_parts = []

        for run in runs:
            text = run["text"]
            if not text:
                continue

            # Apply text formatting per run
            if run.get("bold") and run.get("italic"):
                text = f"***{text}***"
            elif run.get("bold"):
                text = f"**{text}**"
            elif run.get("italic"):
                text = f"*{text}*"

            # Apply hyperlink per run
            if run.get("hyperlink"):
                text = f"[{text}]({run['hyperlink']})"

            formatted_parts.append(text)

        return "".join(formatted_parts)

    # Keep all the other conversion methods unchanged...
    def _convert_table_to_markdown(self, block):
        """Convert table data to standard markdown table format."""
        if not block["data"]:
            return ""

        markdown = ""
        for i, row in enumerate(block["data"]):
            # Escape pipes in cell content to prevent table structure breaks
            escaped_row = [cell.replace("|", "\\|") for cell in row]
            markdown += "| " + " | ".join(escaped_row) + " |\n"

            # Add separator after header row (first row)
            if i == 0:
                markdown += "| " + " | ".join("---" for _ in row) + " |\n"

        return markdown

    def _convert_image_to_markdown(self, block):
        """Convert image blocks to markdown image syntax."""
        # Basic image markdown syntax
        image_md = f"![{block['alt_text']}](image)"

        # Wrap in hyperlink if present
        if block.get("hyperlink"):
            image_md = f"[{image_md}]({block['hyperlink']})"

        return image_md

    def _convert_chart_to_markdown(self, block):
        """Convert chart blocks to markdown with diagram candidate annotation."""
        # Basic chart information
        chart_md = f"**Chart: {block.get('title', 'Untitled Chart')}**\n"
        chart_md += f"*Chart Type: {block.get('chart_type', 'unknown')}*\n\n"

        # Add data summary if available
        if block.get('categories') and block.get('series'):
            chart_md += "Data:\n"
            for series in block['series']:
                if series.get('name'):
                    chart_md += f"- {series['name']}: "
                    if series.get('values'):
                        # Show first 5 values with truncation indicator
                        chart_md += ", ".join(map(str, series['values'][:5]))
                        if len(series['values']) > 5:
                            chart_md += "..."
                    chart_md += "\n"

        # Add diagram candidate annotation for downstream processing
        chart_md += f"\n<!-- DIAGRAM_CANDIDATE: chart, type={block.get('chart_type', 'unknown')} -->\n"

        # Apply hyperlink if present
        if block.get("hyperlink"):
            chart_md = f"[{chart_md}]({block['hyperlink']})"

        return chart_md

    # REMOVED: All the old title detection methods that used pattern matching
    # - _convert_slide_titles_to_headings (no longer needed)
    # - _is_likely_slide_title (no longer needed)
    # - _extract_title_from_bullet (no longer needed)
    # - _contains_multiple_sentences (no longer needed)

    # The XML semantic analysis already identified titles correctly

