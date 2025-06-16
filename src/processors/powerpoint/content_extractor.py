"""
Content Extractor - Updated to pass through XML semantic role information
Now captures semantic roles (title, subtitle, content, other) for downstream processing
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET


class ContentExtractor:
    """
    Extracts content from various PowerPoint shape types with semantic role preservation.
    """

    def extract_shape_content(self, shape, text_processor, accessibility_extractor=None):
        """
        Main extraction router - delegates based on shape type and captures semantic role.

        NEW: Now captures semantic role from XML analysis for proper title detection.

        Args:
            shape: python-pptx Shape object
            text_processor: TextProcessor instance for text handling
            accessibility_extractor: AccessibilityOrderExtractor for semantic role detection

        Returns:
            dict: Content block with semantic role information
        """
        # Capture basic shape info for diagram analysis
        shape_info = self._get_shape_analysis_info(shape)

        # NEW: Capture semantic role from XML analysis
        semantic_role = "other"
        if accessibility_extractor:
            semantic_role = accessibility_extractor._get_semantic_role_from_xml(shape)

        content_block = None

        try:
            # Route based on shape type using explicit type checking
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                content_block = self.extract_image(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                content_block = self.extract_table(shape.table, text_processor)
            elif hasattr(shape, 'has_chart') and shape.has_chart:
                content_block = self.extract_chart(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                content_block = self.extract_group(shape, text_processor, accessibility_extractor)
            elif hasattr(shape, 'text_frame') and shape.text_frame:
                content_block = text_processor.extract_text_frame(shape.text_frame, shape)
            elif hasattr(shape, 'text') and shape.text:
                content_block = text_processor.extract_plain_text(shape)
        except Exception as e:
            print(f"Warning: Error extracting shape content: {e}")
            return None

        # Handle shapes without text content (for diagram analysis)
        if not content_block:
            content_block = self._create_non_text_content_block(shape, shape_info)

        # Add shape analysis info and semantic role for downstream processing
        if content_block:
            try:
                content_block.update(shape_info)
                # NEW: Add semantic role information
                content_block["semantic_role"] = semantic_role
            except Exception as e:
                print(f"Warning: Error adding shape info: {e}")

        return content_block

    def extract_group(self, shape, text_processor, accessibility_extractor=None):
        """
        Extract content from grouped shapes using recursive processing.
        Updated to pass through accessibility_extractor for semantic role detection.
        """
        try:
            extracted_blocks = []

            for child_shape in shape.shapes:
                # Apply same extraction logic to each child shape
                if hasattr(child_shape, 'text_frame') and child_shape.text_frame:
                    text_block = text_processor.extract_text_frame(child_shape.text_frame, child_shape)
                    if text_block:
                        # Add semantic role for group children
                        if accessibility_extractor:
                            semantic_role = accessibility_extractor._get_semantic_role_from_xml(child_shape)
                            text_block["semantic_role"] = semantic_role
                        extracted_blocks.append(text_block)
                elif hasattr(child_shape, 'text') and child_shape.text:
                    text_block = text_processor.extract_plain_text(child_shape)
                    if text_block:
                        # Add semantic role for group children
                        if accessibility_extractor:
                            semantic_role = accessibility_extractor._get_semantic_role_from_xml(child_shape)
                            text_block["semantic_role"] = semantic_role
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
                    nested_group = self.extract_group(child_shape, text_processor, accessibility_extractor)
                    if nested_group and nested_group.get("extracted_blocks"):
                        # Flatten nested group content into current level
                        extracted_blocks.extend(nested_group["extracted_blocks"])

            # Return group container if any content was extracted
            if extracted_blocks:
                return {
                    "type": "group",
                    "extracted_blocks": extracted_blocks,
                    "hyperlink": self._extract_shape_hyperlink(shape),
                    "semantic_role": "group"  # Groups don't have semantic roles themselves
                }

            return None

        except Exception as e:
            print(f"Error extracting group: {e}")
            return None

    # Keep all other existing methods unchanged...
    def extract_image(self, shape):
        """Extract image information with comprehensive alt text detection."""
        alt_text = "Image"

        try:
            if hasattr(shape, 'alt_text') and shape.alt_text:
                alt_text = shape.alt_text
            elif hasattr(shape, 'image') and hasattr(shape.image, 'alt_text') and shape.image.alt_text:
                alt_text = shape.image.alt_text
            elif hasattr(shape, '_element'):
                try:
                    xml_str = str(shape._element.xml) if hasattr(shape._element, 'xml') else ""
                    if xml_str:
                        root = ET.fromstring(xml_str)
                        for elem in root.iter():
                            if 'descr' in elem.attrib and elem.attrib['descr']:
                                alt_text = elem.attrib['descr']
                                break
                            elif 'title' in elem.attrib and elem.attrib['title']:
                                alt_text = elem.attrib['title']
                                break
                except:
                    pass
        except:
            pass

        return {
            "type": "image",
            "alt_text": alt_text.strip() if alt_text else "Image",
            "hyperlink": self._extract_shape_hyperlink(shape)
        }

    def extract_table(self, table, text_processor):
        """Extract table data with cell-level text processing."""
        if not table.rows:
            return None

        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_content = ""

                if hasattr(cell, 'text_frame') and cell.text_frame:
                    cell_paras = []
                    for para in cell.text_frame.paragraphs:
                        if para.text.strip():
                            para_processed = text_processor.process_paragraph(para)
                            if para_processed and para_processed['hints']['is_bullet']:
                                level = para_processed['hints']['bullet_level']
                                indent = "  " * level
                                cell_paras.append(f"{indent}• {para_processed['clean_text']}")
                            elif para_processed:
                                cell_paras.append(para_processed['clean_text'])
                    cell_content = " ".join(cell_paras)
                else:
                    cell_content = cell.text.strip() if hasattr(cell, 'text') else ""

                row_data.append(cell_content)
            table_data.append(row_data)

        return {
            "type": "table",
            "data": table_data
        }

    def extract_chart(self, shape):
        """Extract chart/diagram information for potential Mermaid conversion."""
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

            try:
                if hasattr(chart, 'chart_title') and chart.chart_title and hasattr(chart.chart_title, 'text_frame'):
                    chart_data["title"] = chart.chart_title.text_frame.text.strip()
            except:
                pass

            try:
                if hasattr(chart, 'plots') and chart.plots:
                    plot = chart.plots[0]

                    if hasattr(plot, 'categories') and plot.categories:
                        chart_data["categories"] = [cat.label for cat in plot.categories if hasattr(cat, 'label')]

                    if hasattr(plot, 'series') and plot.series:
                        for series in plot.series:
                            series_data = {
                                "name": series.name if hasattr(series, 'name') else "",
                                "values": []
                            }
                            if hasattr(series, 'values'):
                                try:
                                    series_data["values"] = [val for val in series.values if val is not None]
                                except:
                                    pass
                            chart_data["series"].append(series_data)
            except:
                pass

            return chart_data

        except Exception:
            return {
                "type": "chart",
                "chart_type": "unknown",
                "title": "Chart",
                "data_points": [],
                "categories": [],
                "series": [],
                "hyperlink": self._extract_shape_hyperlink(shape)
            }

    def _create_non_text_content_block(self, shape, shape_info):
        """Create content blocks for shapes without text content."""
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.LINE:
                return {"type": "line", "line_type": "simple"}
            elif shape.shape_type == MSO_SHAPE_TYPE.CONNECTOR:
                return {"type": "line", "line_type": "connector"}
            elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
                return {"type": "line", "line_type": "freeform"}
            elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                if self._is_arrow_shape(shape_info["auto_shape_type"]):
                    return {"type": "arrow", "arrow_type": shape_info["auto_shape_type"]}
                else:
                    return {"type": "shape", "shape_subtype": "auto_shape"}
            else:
                return {"type": "shape", "shape_subtype": "generic"}
        except Exception:
            return {"type": "shape", "shape_subtype": "unknown"}

    def _get_shape_analysis_info(self, shape):
        """Get basic shape information for diagram analysis and debugging."""
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

        try:
            if hasattr(shape, 'shape_type'):
                shape_info["shape_type"] = str(shape.shape_type).split('.')[-1]
        except:
            shape_info["shape_type"] = "unknown"

        try:
            if hasattr(shape, 'auto_shape_type'):
                shape_info["auto_shape_type"] = str(shape.auto_shape_type).split('.')[-1]
        except:
            pass

        return shape_info

    def _is_arrow_shape(self, auto_shape_type):
        """Determine if an auto shape is an arrow type."""
        if not auto_shape_type:
            return False

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
        """Extract shape-level hyperlinks with URL normalization."""
        try:
            if hasattr(shape, 'click_action') and shape.click_action:
                if hasattr(shape.click_action, 'hyperlink') and shape.click_action.hyperlink:
                    if shape.click_action.hyperlink.address:
                        return self._fix_url(shape.click_action.hyperlink.address)
        except:
            pass
        return None

    def _fix_url(self, url):
        """Normalize URLs to handle common PowerPoint URL formatting issues."""
        if not url:
            return url

        if '@' in url and not url.startswith('mailto:'):
            return f"mailto:{url}"

        if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
            if url.startswith('www.') or any(
                    domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                return f"https://{url}"

        return url

