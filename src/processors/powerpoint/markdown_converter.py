"""
Markdown Converter - Converts structured PowerPoint data to markdown format
Handles text formatting, tables, images, charts, and slide title conversion
Fixed: Bold detection for consistently formatted text

ARCHITECTURE OVERVIEW:
This component converts the structured data from ContentExtractor into clean,
readable markdown format. It handles the final formatting stage of the
processing pipeline and includes intelligent post-processing features.

CONVERSION STRATEGY:
- Type-based routing: Different content types handled by specialized methods
- Formatting preservation: Bold, italic, hyperlinks maintained in markdown
- Structure enhancement: Bullet points, tables, headings properly formatted
- Post-processing: Slide titles converted from bullets to H1 headings

MARKDOWN COMPLIANCE:
- Standard markdown syntax for maximum compatibility
- Proper escaping of special characters (pipes in tables)
- Clean formatting without HTML fallbacks
- Accessible structure with semantic headings

FORMATTING FIXES:
The v2 system fixes bold formatting issues where consistently formatted
text was getting double-formatting (e.g., "**text**" became "**text****").
Now applies formatting once to entire consistently-formatted blocks.
"""

import re


class MarkdownConverter:
    """
    Converts structured presentation data to clean markdown format.

    COMPONENT RESPONSIBILITIES:
    - Convert structured content blocks to markdown syntax
    - Preserve text formatting (bold, italic, hyperlinks)
    - Handle complex structures (tables, lists, groups)
    - Post-process for slide title detection and conversion
    - Generate clean, readable output for downstream processing

    CONVERSION PIPELINE:
    1. Process content blocks by type (text, table, image, etc.)
    2. Apply markdown formatting to individual elements
    3. Combine elements with proper spacing
    4. Post-process to enhance structure (slide titles)
    5. Return final markdown document

    FORMATTING PHILOSOPHY:
    - Semantic structure over visual formatting
    - Clean, readable output for human review
    - Machine-parseable for further processing
    - Accessibility-conscious heading hierarchy
    """

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """
        Main conversion entry point - processes entire presentation.

        PROCESSING ALGORITHM:
        1. Iterate through all slides in structured data
        2. Add slide markers for debugging and structure
        3. Process each content block with type-specific handlers
        4. Combine results with proper spacing
        5. Apply post-processing (slide title conversion)

        SLIDE MARKERS:
        HTML comments mark slide boundaries for debugging and
        post-processing. Format: <!-- Slide N --> where N is slide number.

        SPACING STRATEGY:
        Double newlines between major elements provide clean separation
        while maintaining readability. Empty elements are filtered out.

        TITLE CONVERSION:
        Optional post-processing converts first bullet point on each
        slide to H1 heading if it appears to be a slide title.

        Args:
            data (dict): Structured presentation data from ContentExtractor
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1

        Returns:
            str: Complete markdown document with all slides
        """
        markdown_parts = []

        for slide in data["slides"]:
            # Add slide marker for debugging and post-processing
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            # Process each content block with type-specific handling
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

        # Combine all parts with proper spacing
        markdown_content = "\n\n".join(filter(None, markdown_parts))

        # Apply post-processing enhancements
        if convert_slide_titles:
            markdown_content = self._convert_slide_titles_to_headings(markdown_content)

        return markdown_content

    def _convert_text_block_to_markdown(self, block):
        """
        Convert text content blocks to markdown with formatting.

        TEXT BLOCK PROCESSING:
        1. Process each paragraph individually
        2. Apply paragraph-level formatting (bullets, headings)
        3. Combine paragraphs with newline separation
        4. Apply block-level hyperlinks if present

        PARAGRAPH ORDERING:
        Paragraphs are processed in the order provided by the
        AccessibilityOrderExtractor to maintain proper reading flow.

        HYPERLINK HANDLING:
        Shape-level hyperlinks make the entire text block clickable.
        Applied as outermost wrapper around all content.

        Args:
            block (dict): Text content block with paragraphs

        Returns:
            str: Formatted markdown text
        """
        lines = []

        # Process each paragraph individually
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

        PROCESSING PIPELINE:
        1. Build formatted text from runs (handles bold/italic/hyperlinks)
        2. Apply structural formatting based on paragraph hints
        3. Return properly formatted markdown paragraph

        STRUCTURAL FORMATTING HIERARCHY:
        1. Bullets: Convert to markdown list items with indentation
        2. Numbered lists: Convert to markdown numbered lists
        3. Headings: Convert to markdown headings (## or ###)
        4. Regular text: Return with inline formatting only

        HINT SYSTEM:
        The TextProcessor provides hints about paragraph characteristics:
        - is_bullet: Paragraph should be formatted as bullet point
        - is_numbered: Paragraph should be numbered list item
        - likely_heading: Paragraph appears to be a heading
        - bullet_level: Indentation level for nested bullets

        Args:
            para (dict): Paragraph data with text, formatting, and hints

        Returns:
            str: Formatted markdown paragraph
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

        # Handle headings with appropriate level
        elif hints.get("likely_heading", False):
            # Determine heading level based on content characteristics
            if hints.get("all_caps") or len(para["clean_text"]) < 30:
                return f"## {formatted_text}"  # H2 for short/caps headings
            else:
                return f"### {formatted_text}"  # H3 for longer headings

        # Regular paragraph
        else:
            return formatted_text

    def _build_formatted_text_from_runs(self, runs, clean_text):
        """
        Build formatted text from runs with FIXED consistent formatting handling.

        FORMATTING STRATEGY EVOLUTION:
        Previous versions applied formatting per-run, causing issues like:
        "**Read this ****fifth**" for consistently bold text.

        NEW APPROACH:
        1. Check if ALL runs have identical formatting
        2. If consistent: Apply formatting once to entire text
        3. If mixed: Apply formatting per-run as before

        CONSISTENCY DETECTION:
        - All bold: Every run has bold=True
        - All italic: Every run has italic=True
        - All hyperlinked: Every run has same hyperlink URL
        - Mixed formatting: Fall back to per-run processing

        HYPERLINK PRECEDENCE:
        When all runs share the same hyperlink, that takes precedence
        over consistent text formatting to avoid nested link syntax.

        Args:
            runs (list): List of formatted text runs
            clean_text (str): Clean text content for consistent formatting

        Returns:
            str: Properly formatted markdown text
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

    def _convert_table_to_markdown(self, block):
        """
        Convert table data to standard markdown table format.

        MARKDOWN TABLE SPECIFICATION:
        - Pipe-delimited cells: | Cell 1 | Cell 2 |
        - Header separator: | --- | --- | after first row
        - Proper escaping: \| for literal pipe characters

        PROCESSING ALGORITHM:
        1. Iterate through table rows
        2. Escape pipe characters in cell content
        3. Format as pipe-delimited markdown
        4. Add header separator after first row

        HEADER DETECTION:
        First row is automatically treated as header and gets
        separator line. This is standard markdown table behavior.

        ESCAPING STRATEGY:
        Pipe characters in cell content are escaped as \| to
        prevent breaking table structure.

        Args:
            block (dict): Table content block with data array

        Returns:
            str: Formatted markdown table
        """
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
        """
        Convert image blocks to markdown image syntax.

        MARKDOWN IMAGE SYNTAX:
        ![alt_text](image_url) for standard images
        [![alt_text](image_url)](hyperlink) for linked images

        ALT TEXT IMPORTANCE:
        Alt text is crucial for accessibility and content understanding.
        Always included even if just placeholder text.

        HYPERLINK HANDLING:
        Images can have hyperlinks making them clickable. These are
        handled by wrapping the image syntax in link syntax.

        PLACEHOLDER HANDLING:
        Since we don't extract actual image files, "image" is used
        as placeholder URL. Downstream processing could replace this
        with actual image references.

        Args:
            block (dict): Image content block with alt text and hyperlink

        Returns:
            str: Formatted markdown image
        """
        # Basic image markdown syntax
        image_md = f"![{block['alt_text']}](image)"

        # Wrap in hyperlink if present
        if block.get("hyperlink"):
            image_md = f"[{image_md}]({block['hyperlink']})"

        return image_md

    def _convert_chart_to_markdown(self, block):
        """
        Convert chart blocks to markdown with diagram candidate annotation.

        CHART REPRESENTATION STRATEGY:
        1. Title and type information for human readability
        2. Data summary for content understanding
        3. Diagram candidate comment for downstream processing
        4. Hyperlink handling for interactive charts

        DIAGRAM CANDIDATE ANNOTATION:
        Charts are potential candidates for Mermaid diagram conversion.
        The comment provides metadata for Claude or other processors
        to identify conversion opportunities.

        DATA SERIALIZATION:
        Chart data is summarized in human-readable format:
        - Series names with sample values
        - Truncation for long data series (first 5 values)
        - Clear labeling of chart type and structure

        Args:
            block (dict): Chart content block with metadata and data

        Returns:
            str: Formatted markdown chart representation
        """
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

    def _convert_group_to_markdown(self, block):
        """
        Convert grouped shapes to markdown by processing extracted content.

        GROUP PROCESSING STRATEGY:
        1. Process each extracted block from the group
        2. Apply same conversion logic as top-level content
        3. Handle invisible elements (lines, arrows) appropriately
        4. Combine results with proper spacing

        CONTENT TYPE HANDLING:
        - Text: Full text conversion with formatting
        - Images: Standard image markdown
        - Tables: Full table conversion
        - Charts: Chart representation with metadata
        - Lines/Arrows: Skip (invisible but tracked for diagram analysis)
        - Shapes: Basic shape representation

        INVISIBLE ELEMENTS:
        Lines and arrows don't produce visible markdown content but
        are important for diagram analysis. They're processed by
        ContentExtractor but don't generate markdown output.

        Args:
            block (dict): Group content block with extracted_blocks

        Returns:
            str: Formatted markdown representation of group content
        """
        extracted_blocks = block.get("extracted_blocks", [])

        if not extracted_blocks:
            return ""

        # Process each extracted block with type-specific handling
        content_parts = []

        for extracted_block in extracted_blocks:
            content = None

            if extracted_block["type"] == "text":
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

    def _convert_slide_titles_to_headings(self, markdown_content):
        """
        Post-process markdown to convert slide titles from bullets to H1 headings.

        TITLE DETECTION STRATEGY:
        1. Find slide markers (<!-- Slide N -->)
        2. Look for first bullet point after each marker
        3. Evaluate if bullet appears to be a slide title
        4. Convert qualifying bullets to H1 headings

        TITLE CHARACTERISTICS:
        Bullets are converted to titles if they:
        - Appear first on the slide (after slide marker)
        - Meet title length and format criteria
        - Don't have document-style characteristics

        PROCESSING ALGORITHM:
        1. Split content into lines for line-by-line processing
        2. Track slide markers and process following content
        3. Apply title detection to first bullet after each marker
        4. Replace qualifying bullets with H1 syntax

        PRESERVATION:
        Non-title bullets are left unchanged to maintain
        original document structure where appropriate.

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
        Determine if a markdown line is likely a slide title.

        TITLE DETECTION CRITERIA:
        1. Must be a bullet point (starts with '- ')
        2. Reasonable title length (≤150 characters)
        3. No ending punctuation (titles don't end with . ! ? ; :)
        4. Single phrase/sentence (not multiple sentences)
        5. No document-style lead-ins ("The following", "Here are")

        POSITIVE INDICATORS:
        - All caps text (presentation titles often all caps)
        - Title case formatting
        - Short phrases (≤10 words)
        - Title keywords (overview, introduction, agenda, etc.)

        DECISION LOGIC:
        Must meet all basic criteria AND have at least one positive
        indicator to be considered a title candidate.

        Args:
            line (str): Markdown line to evaluate

        Returns:
            bool: True if the line appears to be a slide title
        """
        if not line.strip():
            return False

        # Must be a bullet point to be convertible
        if not line.startswith('- '):
            return False

        # Extract the text content
        text_content = line[2:].strip()

        # Basic title characteristics (all must be true)
        title_indicators = [
            len(text_content) <= 150,  # Reasonable title length
            not text_content.endswith(('.', '!', '?', ';', ':')),  # No ending punctuation
            not self._contains_multiple_sentences(text_content),  # Single phrases
            not text_content.lower().startswith(('the following', 'here are', 'this slide', 'key points')),
        ]

        # Positive indicators (at least one should be true)
        positive_indicators = [
            text_content.isupper(),  # ALL CAPS suggests title
            text_content.istitle(),  # Title Case suggests title
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

        EXTRACTION PROCESS:
        1. Remove bullet prefix ('- ')
        2. Strip whitespace
        3. Remove markdown formatting artifacts
        4. Return clean title text

        ARTIFACT REMOVAL:
        Removes common markdown formatting characters that might
        appear in extracted text but shouldn't be in headings:
        - Asterisks (*) from bold/italic formatting
        - Underscores (_) from italic formatting
        - Backticks (`) from code formatting

        Args:
            bullet_line (str): Bullet point line with '- ' prefix

        Returns:
            str: Clean title text suitable for H1 heading
        """
        # Remove bullet prefix
        title_text = bullet_line[2:].strip()

        # Clean up common title artifacts from markdown formatting
        title_text = title_text.strip('*_`')  # Remove markdown formatting chars

        return title_text

    def _contains_multiple_sentences(self, text):
        """
        Check if text contains multiple sentences using punctuation patterns.

        SENTENCE DETECTION:
        Uses regex to find sentence-ending punctuation (. ! ?) followed
        by whitespace and a capital letter, indicating sentence boundaries.

        RATIONALE:
        Titles are typically single phrases or sentences. Multiple
        sentences suggest document content rather than a title.

        PATTERN: [.!?]\s+[A-Z]
        - [.!?]: Sentence-ending punctuation
        - \s+: One or more whitespace characters
        - [A-Z]: Capital letter starting next sentence

        Args:
            text (str): Text to analyze

        Returns:
            bool: True if text appears to contain multiple sentences
        """
        # Pattern for sentence boundaries: punctuation + space + capital letter
        sentence_pattern = r'[.!?]\s+[A-Z]'
        return bool(re.search(sentence_pattern, text))