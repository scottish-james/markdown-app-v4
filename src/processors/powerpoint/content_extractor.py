"""
Content Extractor - Handles extracting content from different PowerPoint shape types
Specializes in extracting text, images, tables, charts, and grouped shapes

ARCHITECTURE OVERVIEW:
This component implements a type-based extraction strategy where different
PowerPoint shape types are handled by specialized extraction methods. It acts
as a router that delegates to appropriate handlers based on shape type.

SHAPE TYPE HANDLING:
- Text shapes: Delegated to TextProcessor for advanced formatting
- Images: Alt text and hyperlink extraction
- Tables: Cell-by-cell content with formatting preservation
- Charts: Metadata extraction for potential Mermaid conversion
- Groups: Recursive processing of child shapes
- Lines/Arrows: Shape analysis for diagram detection

INTEGRATION POINTS:
- Works closely with TextProcessor for text formatting
- Provides shape analysis data for DiagramAnalyzer
- Handles hyperlinks and accessibility information
- Supports both XML-rich and fallback processing modes

ERROR HANDLING PHILOSOPHY:
- Graceful degradation: extraction failures don't stop processing
- Defensive programming: all shape access wrapped in try/catch
- Fallback content: provides basic alternatives when detailed extraction fails
- Warning logging: alerts developers to extraction issues without crashing
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET


class ContentExtractor:
    """
    Extracts content from various PowerPoint shape types with type-specific handling.

    COMPONENT RESPONSIBILITIES:
    - Route shape processing based on MSO_SHAPE_TYPE
    - Extract content while preserving formatting and metadata
    - Provide shape analysis data for diagram detection
    - Handle nested structures (groups, tables) recursively
    - Extract accessibility information (alt text, hyperlinks)

    PROCESSING PATTERNS:
    - Delegation: Routes to specialized extraction methods
    - Composition: Uses TextProcessor for text formatting
    - Recursion: Handles nested groups and complex structures
    - Fallback: Provides basic content when detailed extraction fails

    OUTPUT FORMAT:
    Returns standardized content blocks with consistent structure:
    {
        "type": "text|image|table|chart|group|shape|line|arrow",
        "content_specific_fields": "...",
        "shape_analysis_info": "..." // For diagram detection
    }
    """

    def extract_shape_content(self, shape, text_processor):
        """
        Main extraction router - delegates based on shape type.

        ROUTING STRATEGY:
        1. Identify shape type using MSO_SHAPE_TYPE enum
        2. Route to appropriate specialized extraction method
        3. Add shape analysis information for diagram detection
        4. Handle extraction failures gracefully with fallbacks

        SHAPE TYPE PRIORITY:
        - Special types first: PICTURE, TABLE, charts
        - Groups: Recursive processing
        - Text: Most common, handled by TextProcessor
        - Fallback: Create basic content blocks for unknown types

        ERROR HANDLING:
        - Individual extraction failures don't stop processing
        - Warnings logged for debugging
        - Fallback content created for failed extractions
        - Shape analysis info always added when possible

        Args:
            shape: python-pptx Shape object
            text_processor: TextProcessor instance for text handling

        Returns:
            dict: Content block or None if no extractable content
        """
        # Capture basic shape info for diagram analysis
        # This is done first to ensure we have shape data even if extraction fails
        shape_info = self._get_shape_analysis_info(shape)

        content_block = None

        try:
            # Route based on shape type using explicit type checking
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                content_block = self.extract_image(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                content_block = self.extract_table(shape.table, text_processor)
            elif hasattr(shape, 'has_chart') and shape.has_chart:
                # Chart detection: some shapes have charts embedded
                content_block = self.extract_chart(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                # Groups require recursive processing of children
                content_block = self.extract_group(shape, text_processor)
            elif hasattr(shape, 'text_frame') and shape.text_frame:
                # Text shapes with full text frame structure
                content_block = text_processor.extract_text_frame(shape.text_frame, shape)
            elif hasattr(shape, 'text') and shape.text:
                # Simple text shapes without full text frame
                content_block = text_processor.extract_plain_text(shape)
        except Exception as e:
            print(f"Warning: Error extracting shape content: {e}")
            return None

        # Handle shapes without text content (for diagram analysis)
        if not content_block:
            content_block = self._create_non_text_content_block(shape, shape_info)

        # Add shape analysis info for diagram detection
        if content_block:
            try:
                content_block.update(shape_info)
            except Exception as e:
                print(f"Warning: Error adding shape info: {e}")

        return content_block

    def extract_image(self, shape):
        """
        Extract image information with comprehensive alt text detection.

        ALT TEXT EXTRACTION STRATEGY:
        1. Try direct alt_text attribute (most reliable)
        2. Try image.alt_text property (alternative accessor)
        3. Parse XML for description/title attributes (fallback)
        4. Default to "Image" if all methods fail

        XML ALT TEXT PARSING:
        PowerPoint stores alt text in various XML attributes:
        - descr: Description attribute (primary)
        - title: Title attribute (secondary)
        - Various namespace-specific attributes

        ACCESSIBILITY IMPORTANCE:
        Alt text is crucial for accessibility and content understanding.
        Multiple extraction methods ensure we capture this information
        even when primary accessors fail.

        Args:
            shape: python-pptx Picture shape

        Returns:
            dict: Image content block with alt text and hyperlink
        """
        alt_text = "Image"

        try:
            # Method 1: Direct alt_text attribute (most common)
            if hasattr(shape, 'alt_text') and shape.alt_text:
                alt_text = shape.alt_text
            # Method 2: Image object alt_text property
            elif hasattr(shape, 'image') and hasattr(shape.image, 'alt_text') and shape.image.alt_text:
                alt_text = shape.image.alt_text
            # Method 3: XML extraction (fallback for edge cases)
            elif hasattr(shape, '_element'):
                try:
                    xml_str = str(shape._element.xml) if hasattr(shape._element, 'xml') else ""
                    if xml_str:
                        root = ET.fromstring(xml_str)
                        # Search for alt text in XML attributes
                        for elem in root.iter():
                            if 'descr' in elem.attrib and elem.attrib['descr']:
                                alt_text = elem.attrib['descr']
                                break
                            elif 'title' in elem.attrib and elem.attrib['title']:
                                alt_text = elem.attrib['title']
                                break
                except:
                    # XML parsing can fail - continue with default
                    pass
        except:
            # All alt text extraction failed - use default
            pass

        return {
            "type": "image",
            "alt_text": alt_text.strip() if alt_text else "Image",
            "hyperlink": self._extract_shape_hyperlink(shape)
        }

    def extract_table(self, table, text_processor):
        """
        Extract table data with cell-level text processing.

        TABLE PROCESSING STRATEGY:
        1. Iterate through rows and cells systematically
        2. Process each cell's text content using TextProcessor
        3. Handle bullet points within cells with proper indentation
        4. Preserve text formatting while making it markdown-compatible

        CELL TEXT PROCESSING:
        - Uses TextProcessor for consistent bullet/formatting handling
        - Converts bullets to markdown format with proper indentation
        - Handles multi-paragraph cells by joining content
        - Maintains text formatting across cell boundaries

        BULLET HANDLING IN CELLS:
        PowerPoint tables can contain bullet points within cells.
        These are converted to markdown-style bullets with indentation
        to preserve visual hierarchy.

        ERROR HANDLING:
        - Empty tables return None rather than empty structure
        - Cell access failures default to empty string
        - Text processing failures fall back to plain text

        Args:
            table: python-pptx Table object
            text_processor: TextProcessor for handling cell text

        Returns:
            dict: Table content block with processed cell data
        """
        if not table.rows:
            return None

        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_content = ""

                # Process cell text with full TextProcessor capabilities
                if hasattr(cell, 'text_frame') and cell.text_frame:
                    cell_paras = []
                    for para in cell.text_frame.paragraphs:
                        if para.text.strip():
                            # Process paragraph for bullets and formatting
                            para_processed = text_processor.process_paragraph(para)
                            if para_processed and para_processed['hints']['is_bullet']:
                                # Convert bullets to markdown format with indentation
                                level = para_processed['hints']['bullet_level']
                                indent = "  " * level
                                cell_paras.append(f"{indent}• {para_processed['clean_text']}")
                            elif para_processed:
                                # Regular paragraph content
                                cell_paras.append(para_processed['clean_text'])
                    cell_content = " ".join(cell_paras)
                else:
                    # Fallback: Simple text extraction
                    cell_content = cell.text.strip() if hasattr(cell, 'text') else ""

                row_data.append(cell_content)
            table_data.append(row_data)

        return {
            "type": "table",
            "data": table_data
        }

    def extract_chart(self, shape):
        """
        Extract chart/diagram information for potential Mermaid conversion.

        CHART ANALYSIS PURPOSE:
        Charts are potential candidates for Mermaid diagram conversion.
        This method extracts metadata and data structure needed for
        intelligent diagram generation.

        EXTRACTION STRATEGY:
        1. Chart type identification for Mermaid compatibility assessment
        2. Title extraction for diagram labeling
        3. Data series and categories for structure analysis
        4. Error handling for corrupted or unsupported chart types

        MERMAID CONVERSION POTENTIAL:
        - Flowcharts: Process flows and decision trees
        - Organization charts: Hierarchy relationships
        - Sequence diagrams: Timeline-based interactions
        - Pie charts: Proportional data visualization

        DATA STRUCTURE EXTRACTION:
        - Categories: X-axis labels or grouping information
        - Series: Data sets with names and values
        - Relationships: Implicit in chart structure

        Args:
            shape: python-pptx Shape with chart

        Returns:
            dict: Chart content block with structure for Mermaid analysis
        """
        try:
            chart = shape.chart
            chart_data = {
                "type": "chart",
                "chart_type": str(chart.chart_type) if hasattr(chart, 'chart_type') else "unknown",
                "title": "",
                "data_points": [],
                "categories": [],
                "series": [],
                "hyperlink": self._extract_shape_hyperlink(shape)
            }

            # Extract chart title with multiple access patterns
            try:
                if hasattr(chart, 'chart_title') and chart.chart_title and hasattr(chart.chart_title, 'text_frame'):
                    chart_data["title"] = chart.chart_title.text_frame.text.strip()
            except:
                # Chart title access can fail
                pass

            # Extract data structure for potential Mermaid conversion
            try:
                if hasattr(chart, 'plots') and chart.plots:
                    plot = chart.plots[0]  # Use first plot for analysis

                    # Extract categories (X-axis labels, groupings)
                    if hasattr(plot, 'categories') and plot.categories:
                        chart_data["categories"] = [cat.label for cat in plot.categories if hasattr(cat, 'label')]

                    # Extract data series (datasets with values)
                    if hasattr(plot, 'series') and plot.series:
                        for series in plot.series:
                            series_data = {
                                "name": series.name if hasattr(series, 'name') else "",
                                "values": []
                            }
                            if hasattr(series, 'values'):
                                try:
                                    # Extract numeric values, filtering out None
                                    series_data["values"] = [val for val in series.values if val is not None]
                                except:
                                    # Value extraction can fail for complex chart types
                                    pass
                            chart_data["series"].append(series_data)
            except:
                # Data extraction can fail for unsupported chart types
                pass

            return chart_data

        except Exception:
            # Fallback for charts we can't parse at all
            return {
                "type": "chart",
                "chart_type": "unknown",
                "title": "Chart",
                "data_points": [],
                "categories": [],
                "series": [],
                "hyperlink": self._extract_shape_hyperlink(shape)
            }

    def extract_group(self, shape, text_processor):
        """
        Extract content from grouped shapes using recursive processing.

        GROUP PROCESSING STRATEGY:
        1. Iterate through all child shapes in the group
        2. Apply same extraction logic recursively to each child
        3. Handle nested groups with additional recursion
        4. Collect all extracted content into a group block

        RECURSION HANDLING:
        Groups can contain other groups (nested grouping). This method
        handles arbitrary nesting depth by recursively calling itself
        for nested groups and flattening the results.

        CHILD SHAPE PROCESSING:
        Each child shape is processed using the same type-based routing
        as top-level shapes, ensuring consistent extraction regardless
        of grouping structure.

        CONTENT AGGREGATION:
        All extracted content blocks from children are collected into
        a single group container, preserving the hierarchical structure
        while making content accessible.

        Args:
            shape: python-pptx Group shape
            text_processor: TextProcessor for text handling

        Returns:
            dict: Group content block with all extracted child content
        """
        try:
            extracted_blocks = []

            for child_shape in shape.shapes:
                # Apply same extraction logic to each child shape
                if hasattr(child_shape, 'text_frame') and child_shape.text_frame:
                    text_block = text_processor.extract_text_frame(child_shape.text_frame, child_shape)
                    if text_block:
                        extracted_blocks.append(text_block)
                elif hasattr(child_shape, 'text') and child_shape.text:
                    text_block = text_processor.extract_plain_text(child_shape)
                    if text_block:
                        extracted_blocks.append(text_block)
                elif child_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image_block = self.extract_image(child_shape)
                    if image_block:
                        extracted_blocks.append(image_block)
                elif child_shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    table_block = self.extract_table(child_shape.table, text_processor)
                    if table_block:
                        extracted_blocks.append(table_block)
                elif hasattr(child_shape, 'has_chart') and child_shape.has_chart:
                    chart_block = self.extract_chart(child_shape)
                    if chart_block:
                        extracted_blocks.append(chart_block)
                elif child_shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    # Handle nested groups recursively
                    nested_group = self.extract_group(child_shape, text_processor)
                    if nested_group and nested_group.get("extracted_blocks"):
                        # Flatten nested group content into current level
                        extracted_blocks.extend(nested_group["extracted_blocks"])

            # Return group container if any content was extracted
            if extracted_blocks:
                return {
                    "type": "group",
                    "extracted_blocks": extracted_blocks,
                    "hyperlink": self._extract_shape_hyperlink(shape)
                }

            return None

        except Exception as e:
            print(f"Error extracting group: {e}")
            return None

    def _create_non_text_content_block(self, shape, shape_info):
        """
        Create content blocks for shapes without text content.

        PURPOSE: Shapes like lines, arrows, and geometric shapes don't contain
        text but are important for diagram analysis. This method creates
        basic content blocks that provide metadata for diagram detection.

        SHAPE CLASSIFICATION:
        - Lines: Simple lines, connectors, freeform paths
        - Arrows: Directional indicators (various arrow types)
        - Auto shapes: PowerPoint's built-in shapes
        - Generic shapes: Fallback for unknown shape types

        DIAGRAM ANALYSIS SUPPORT:
        These content blocks provide the DiagramAnalyzer with information
        about non-text elements that indicate diagram presence (lines
        connecting shapes, directional flow indicators, etc.).

        Args:
            shape: python-pptx Shape object
            shape_info: Shape analysis information

        Returns:
            dict: Content block for non-text shapes
        """
        try:
            # Classify shape based on MSO_SHAPE_TYPE
            if shape.shape_type == MSO_SHAPE_TYPE.LINE:
                return {"type": "line", "line_type": "simple"}
            elif shape.shape_type == MSO_SHAPE_TYPE.CONNECTOR:
                return {"type": "line", "line_type": "connector"}
            elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
                return {"type": "line", "line_type": "freeform"}
            elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                # Auto shapes include arrows and geometric shapes
                if self._is_arrow_shape(shape_info["auto_shape_type"]):
                    return {"type": "arrow", "arrow_type": shape_info["auto_shape_type"]}
                else:
                    return {"type": "shape", "shape_subtype": "auto_shape"}
            else:
                return {"type": "shape", "shape_subtype": "generic"}
        except Exception:
            # Shape type detection failed - use generic fallback
            return {"type": "shape", "shape_subtype": "unknown"}

    def _get_shape_analysis_info(self, shape):
        """
        Get basic shape information for diagram analysis and debugging.

        ANALYSIS INFORMATION:
        - Shape type: MSO_SHAPE_TYPE for classification
        - Auto shape type: Specific type for arrows and special shapes
        - Position: Location and size for spatial analysis
        - This data supports diagram detection algorithms

        POSITION DATA USAGE:
        The DiagramAnalyzer uses position information to:
        - Determine spatial layout patterns
        - Calculate shape distribution and alignment
        - Identify grid-like arrangements
        - Assess visual complexity

        ERROR HANDLING:
        All attribute access is defensive to handle:
        - Missing attributes on some shape types
        - Corrupted shape data
        - Version differences in python-pptx

        Args:
            shape: python-pptx Shape object

        Returns:
            dict: Shape analysis information for diagram detection
        """
        shape_info = {
            "shape_type": "unknown",
            "auto_shape_type": None,
            "position": {
                "top": getattr(shape, 'top', 0),
                "left": getattr(shape, 'left', 0),
                "width": getattr(shape, 'width', 0),
                "height": getattr(shape, 'height', 0)
            }
        }

        # Extract shape type with error handling
        try:
            if hasattr(shape, 'shape_type'):
                # Convert enum to string for serialization
                shape_info["shape_type"] = str(shape.shape_type).split('.')[-1]
        except:
            shape_info["shape_type"] = "unknown"

        # Extract auto shape type for arrow detection
        try:
            if hasattr(shape, 'auto_shape_type'):
                # Auto shape type provides specific shape classification
                shape_info["auto_shape_type"] = str(shape.auto_shape_type).split('.')[-1]
        except:
            # Auto shape type not available for all shape types
            pass

        return shape_info

    def _is_arrow_shape(self, auto_shape_type):
        """
        Determine if an auto shape is an arrow type.

        ARROW DETECTION PURPOSE:
        Arrows are strong indicators of diagram content. This method
        identifies various arrow types from PowerPoint's auto shape
        collection for diagram analysis.

        ARROW TYPE COVERAGE:
        - Directional arrows: left, right, up, down
        - Multi-directional: left-right, up-down, quad
        - Curved arrows: various curved directions
        - Special arrows: bent, U-turn, notched, striped
        - Block arrows: filled arrow shapes

        MAINTENANCE NOTE:
        This list should be updated if PowerPoint adds new arrow types
        or if diagram analysis requires detection of additional arrow
        variants.

        Args:
            auto_shape_type: Auto shape type string from shape analysis

        Returns:
            bool: True if the shape is an arrow type
        """
        if not auto_shape_type:
            return False

        # Comprehensive list of PowerPoint arrow auto shapes
        arrow_types = [
            "LEFT_ARROW", "DOWN_ARROW", "UP_ARROW", "RIGHT_ARROW",
            "LEFT_RIGHT_ARROW", "UP_DOWN_ARROW", "QUAD_ARROW",
            "LEFT_RIGHT_UP_ARROW", "BENT_ARROW", "U_TURN_ARROW",
            "CURVED_LEFT_ARROW", "CURVED_RIGHT_ARROW",
            "CURVED_UP_ARROW", "CURVED_DOWN_ARROW",
            "STRIPED_RIGHT_ARROW", "NOTCHED_RIGHT_ARROW",
            "BLOCK_ARC"
        ]

        return any(arrow_type in auto_shape_type for arrow_type in arrow_types)

    def _extract_shape_hyperlink(self, shape):
        """
        Extract shape-level hyperlinks with URL normalization.

        SHAPE HYPERLINKS:
        PowerPoint allows entire shapes to be hyperlinks, separate from
        text-level hyperlinks within the shape content. This is common
        for images, buttons, and navigation elements.

        ACCESS PATTERN:
        shape.click_action.hyperlink.address is the standard path, but
        each level can be None, requiring defensive navigation.

        URL FIXING:
        Extracted URLs are passed through _fix_url() to handle common
        formatting issues like missing schemes or email addresses
        without mailto: prefixes.

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
            # Hyperlink access can fail in various ways
            pass
        return None

    def _fix_url(self, url):
        """
        Normalize URLs to handle common PowerPoint URL formatting issues.

        COMMON ISSUES FIXED:
        1. Email addresses missing mailto: scheme
        2. Web URLs missing http/https scheme
        3. Relative URLs that need scheme inference

        URL SCHEME DETECTION:
        - Email: Contains @ without existing mailto: scheme
        - Web: Starts with www. or contains common TLDs
        - Already formatted: Has existing scheme - return as-is

        SECURITY CONSIDERATION:
        Defaults to HTTPS for web URLs to encourage secure connections
        and follow modern web standards.

        Args:
            url (str): Potentially malformed URL

        Returns:
            str: Properly formatted URL with appropriate scheme
        """
        if not url:
            return url

        # Fix email addresses missing mailto: scheme
        if '@' in url and not url.startswith('mailto:'):
            return f"mailto:{url}"

        # Fix web URLs missing scheme
        if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
            # Detect web URLs by common patterns
            if url.startswith('www.') or any(
                    domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                return f"https://{url}"  # Default to HTTPS for security

        return url