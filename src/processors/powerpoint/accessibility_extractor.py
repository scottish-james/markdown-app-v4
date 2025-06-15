"""
Simplified Accessibility Order Extractor with XML-first approach
XML-first approach with MarkItDown fallback when XML unavailable
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET
import re


class AccessibilityOrderExtractor:
    """
    Simplified accessibility extractor: XML first, MarkItDown fallback.
    Assumes XML is available for sophisticated processing.
    """

    def __init__(self, use_accessibility_order=True):
        """
        Initialize the accessibility extractor.

        Args:
            use_accessibility_order (bool): Whether to use accessibility order vs positional
        """
        self.use_accessibility_order = use_accessibility_order
        self.last_extraction_method = "not_extracted"

        # XML namespaces for PowerPoint processing
        self.namespaces = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

    def get_slide_reading_order(self, slide, slide_number):
        """
        Get shapes from slide in proper reading order.
        If XML not available, falls back to MarkItDown approach.

        Args:
            slide: python-pptx Slide object
            slide_number (int): Slide number for debugging

        Returns:
            list: Ordered list of shapes in reading order
        """
        if not self.use_accessibility_order:
            # Use simple positional method
            ordered_shapes = self._get_positional_ordered_shapes(slide)
            self.last_extraction_method = "positional_order"
            return ordered_shapes

        # Check if we have XML access
        if not self._has_xml_access(slide):
            # No XML - fall back to MarkItDown approach (simple shape order)
            self.last_extraction_method = "markitdown_fallback"
            return list(slide.shapes)

        try:
            # We have XML - use sophisticated accessibility order extraction
            ordered_shapes = self._get_semantic_accessibility_order(slide)
            self.last_extraction_method = "semantic_accessibility_order"
            return ordered_shapes
        except Exception as e:
            print(f"XML accessibility extraction failed for slide {slide_number}: {e}")
            print("Falling back to simple shape order...")
            self.last_extraction_method = "xml_error_fallback"
            return list(slide.shapes)

    def _has_xml_access(self, slide):
        """
        Check if we have XML access for this slide.

        Args:
            slide: python-pptx Slide object

        Returns:
            bool: True if XML is accessible
        """
        try:
            # Try to access slide XML
            slide_xml = self._get_slide_xml(slide)
            return slide_xml is not None and len(slide_xml) > 0
        except Exception:
            return False

    def _get_semantic_accessibility_order(self, slide):
        """
        Get shapes in semantic accessibility order using XML data.

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes ordered by semantic importance with group children extracted
        """
        # Get all shapes in XML document order first
        xml_ordered_shapes = self._get_xml_document_order(slide)

        # Process each shape, handling groups specially
        final_ordered_shapes = []

        for shape in xml_ordered_shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                # For groups, get internal reading order and add children individually
                group_children = self.get_reading_order_of_grouped_by_shape(shape)
                final_ordered_shapes.extend(group_children)
            else:
                # Regular shape - add directly
                final_ordered_shapes.append(shape)

        # Now separate by semantic importance (titles first, etc.)
        title_shapes = []
        subtitle_shapes = []
        content_shapes = []
        other_shapes = []

        for shape in final_ordered_shapes:
            semantic_role = self._get_semantic_role_from_xml(shape)

            if semantic_role == "title":
                title_shapes.append(shape)
            elif semantic_role == "subtitle":
                subtitle_shapes.append(shape)
            elif semantic_role == "content":
                content_shapes.append(shape)
            else:
                other_shapes.append(shape)

        return title_shapes + subtitle_shapes + content_shapes + other_shapes

    def _get_semantic_role_from_xml(self, shape):
        """
        Get semantic role from XML placeholder types (reliable source).

        Args:
            shape: python-pptx Shape object

        Returns:
            str: Semantic role ('title', 'subtitle', 'content', 'other')
        """
        # Check PowerPoint placeholder types (most reliable)
        try:
            if hasattr(shape, 'placeholder_format') and shape.placeholder_format:
                ph_type = shape.placeholder_format.type
                if hasattr(ph_type, 'name'):
                    ph_name = ph_type.name.upper()

                    # Definitive title detection
                    if any(title_type in ph_name for title_type in ['TITLE', 'CENTER_TITLE']):
                        if 'SUBTITLE' not in ph_name:
                            return "title"

                    # Definitive subtitle detection
                    if 'SUBTITLE' in ph_name:
                        return "subtitle"

                    # Definitive content detection
                    if any(content_type in ph_name for content_type in ['BODY', 'CONTENT', 'TEXT', 'OBJECT']):
                        return "content"
        except Exception:
            pass

        # Fallback: check shape names from XML
        try:
            if hasattr(shape, 'name') and shape.name:
                name_lower = shape.name.lower()

                if any(title_word in name_lower for title_word in ['title', 'heading', 'header']):
                    if 'subtitle' not in name_lower:
                        return "title"

                if 'subtitle' in name_lower:
                    return "subtitle"
        except Exception:
            pass

        # Default classification
        if hasattr(shape, 'text_frame') or hasattr(shape, 'text'):
            return "content"
        else:
            return "other"

    def _get_xml_document_order(self, slide):
        """
        Get shapes in XML document order.

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes in XML document order
        """
        try:
            # Get the slide's XML
            slide_xml = self._get_slide_xml(slide)

            # Parse XML to get shapes in document order
            xml_shape_info = self._parse_slide_xml_for_document_order(slide_xml)

            # Map XML order to python-pptx shapes
            ordered_shapes = self._map_xml_to_pptx_shapes(xml_shape_info, slide.shapes)

            return ordered_shapes
        except Exception as e:
            raise Exception(f"XML document order extraction failed: {e}")

    def _get_positional_ordered_shapes(self, slide):
        """
        Simple positional ordering method (fallback).

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes ordered by position (top-to-bottom, left-to-right)
        """
        positioned_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, 'top') and hasattr(shape, 'left'):
                positioned_shapes.append((shape.top, shape.left, shape))
            else:
                positioned_shapes.append((0, 0, shape))

        positioned_shapes.sort(key=lambda x: (x[0], x[1]))
        return [shape for _, _, shape in positioned_shapes]

    def _parse_slide_xml_for_document_order(self, slide_xml):
        """
        Parse slide XML to extract shapes in document order.

        Args:
            slide_xml (str): Raw XML content of slide

        Returns:
            list: Shape information in document order
        """
        root = ET.fromstring(slide_xml)

        # Find the shape tree containing shapes in document order
        shape_tree = root.find('.//p:spTree', self.namespaces)
        if shape_tree is None:
            raise Exception("No shape tree found in slide XML")

        shape_order_info = []

        # Process all shape elements in exact document order
        for idx, elem in enumerate(shape_tree):
            tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

            # Include all shape types
            if tag_name in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent']:
                shape_info = self._extract_shape_info_from_xml(elem, idx)
                if shape_info:
                    shape_order_info.append(shape_info)

        return shape_order_info

    def _extract_shape_info_from_xml(self, shape_elem, order_index):
        """
        Extract identifying information from a shape's XML element.

        Args:
            shape_elem: XML element for shape
            order_index (int): Position in document order

        Returns:
            dict: Shape identification information
        """
        shape_info = {
            'xml_order': order_index,
            'id': None,
            'name': None,
            'type': shape_elem.tag.split('}')[-1] if '}' in shape_elem.tag else shape_elem.tag,
            'has_text': False,
            'text_content': None
        }

        # Extract ID and name from non-visual properties
        nv_props = (shape_elem.find('.//p:nvSpPr', self.namespaces) or
                    shape_elem.find('.//p:nvPicPr', self.namespaces) or
                    shape_elem.find('.//p:nvGraphicFramePr', self.namespaces) or
                    shape_elem.find('.//p:nvGrpSpPr', self.namespaces) or
                    shape_elem.find('.//p:nvCxnSpPr', self.namespaces))

        if nv_props is not None:
            cnv_pr = nv_props.find('.//p:cNvPr', self.namespaces)
            if cnv_pr is not None:
                shape_info['id'] = cnv_pr.get('id')
                shape_info['name'] = cnv_pr.get('name', '')

        # Check for text content
        text_elements = shape_elem.findall('.//a:t', self.namespaces)
        if text_elements:
            all_text = ' '.join([t.text for t in text_elements if t.text])
            if all_text.strip():
                shape_info['has_text'] = True
                shape_info['text_content'] = all_text.strip()[:50]

        return shape_info

    def _map_xml_to_pptx_shapes(self, xml_shape_info, pptx_shapes):
        """
        Map XML shape information to python-pptx shapes.

        Args:
            xml_shape_info (list): Shape info from XML parsing
            pptx_shapes: python-pptx shapes collection

        Returns:
            list: Ordered python-pptx shapes
        """
        ordered_shapes = []
        used_shapes = set()

        # Create simple lookup by ID and name
        shape_lookup = {}
        for shape in pptx_shapes:
            shape_id = self._get_shape_id(shape)
            shape_name = self._get_shape_name(shape)

            if shape_id:
                shape_lookup[f"id_{shape_id}"] = shape
            if shape_name:
                shape_lookup[f"name_{shape_name}"] = shape

        # Match XML order to shapes
        for xml_info in xml_shape_info:
            matched_shape = None

            # Try ID matching first
            if xml_info['id']:
                matched_shape = shape_lookup.get(f"id_{xml_info['id']}")

            # Try name matching
            if not matched_shape and xml_info['name']:
                matched_shape = shape_lookup.get(f"name_{xml_info['name']}")

            # Add if found and not already used
            if matched_shape and matched_shape not in used_shapes:
                ordered_shapes.append(matched_shape)
                used_shapes.add(matched_shape)

        # Add any remaining shapes
        for shape in pptx_shapes:
            if shape not in used_shapes:
                ordered_shapes.append(shape)

        return ordered_shapes

    def _get_slide_xml(self, slide):
        """Extract raw XML from a slide."""
        try:
            if hasattr(slide, '_element') and hasattr(slide._element, 'xml'):
                return slide._element.xml
            elif hasattr(slide, 'element') and hasattr(slide.element, 'xml'):
                return slide.element.xml
            else:
                slide_part = slide.part if hasattr(slide, 'part') else None
                if slide_part and hasattr(slide_part, '_element'):
                    return slide_part._element.xml
                else:
                    raise Exception("Cannot access slide XML")
        except Exception:
            return None

    def _get_shape_id(self, shape):
        """Extract shape ID from python-pptx shape."""
        try:
            if hasattr(shape, '_element') and hasattr(shape._element, 'xml'):
                xml_str = shape._element.xml
                match = re.search(r'<[^>]*:cNvPr[^>]+id="([^"]+)"', xml_str)
                if match:
                    return match.group(1)
        except:
            pass
        return None

    def _get_shape_name(self, shape):
        """Extract shape name from python-pptx shape."""
        try:
            if hasattr(shape, 'name') and shape.name:
                return shape.name
        except:
            pass
        return None

    def get_last_extraction_method(self):
        """Get the method used in the last extraction."""
        return self.last_extraction_method

    def get_reading_order_of_grouped_by_shape(self, group_shape):
        """
        Extract reading order of shapes within a group using z-axis (stacking order).

        Args:
            group_shape: python-pptx GroupShape object

        Returns:
            list: Child shapes in proper reading order (z-axis based)
        """
        try:
            # First, try XML-based group reading order
            xml_ordered_children = self._get_group_xml_reading_order(group_shape)
            if xml_ordered_children:
                return xml_ordered_children

        except Exception as e:
            print(f"XML group reading order failed: {e}")

        # Fallback: Use z-axis (stacking order)
        return self._get_group_z_axis_order(group_shape)

    def _get_group_xml_reading_order(self, group_shape):
        """
        Extract child shapes from group XML in document order.

        Args:
            group_shape: python-pptx GroupShape object

        Returns:
            list: Child shapes in XML document order
        """
        try:
            # Get the group's XML element
            group_xml = group_shape._element.xml

            # Parse the group XML
            root = ET.fromstring(group_xml)

            # Find child shapes in the group XML
            child_elements = []

            # Look for child shape elements within the group
            for elem in root.iter():
                tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                if tag_name in ['sp', 'pic', 'graphicFrame', 'cxnSp']:
                    # Extract child shape info
                    child_info = self._extract_child_shape_info(elem)
                    if child_info:
                        child_elements.append(child_info)

            # Map XML children to python-pptx child shapes
            if child_elements:
                return self._map_xml_children_to_pptx_children(child_elements, group_shape.shapes)

        except Exception as e:
            print(f"Group XML parsing failed: {e}")

        return None

    def _get_group_z_axis_order(self, group_shape):
        """
        Get child shapes ordered by z-axis (front to back stacking order).

        Args:
            group_shape: python-pptx GroupShape object

        Returns:
            list: Child shapes ordered by z-axis position
        """
        try:
            # Get child shapes with their z-order information
            children_with_z_order = []

            for child_shape in group_shape.shapes:
                z_order = self._get_shape_z_order(child_shape)
                children_with_z_order.append((z_order, child_shape))

            # Sort by z-order (lower values = back, higher values = front)
            children_with_z_order.sort(key=lambda x: x[0])

            # Return just the shapes (without z-order values)
            return [shape for z_order, shape in children_with_z_order]

        except Exception as e:
            print(f"Z-axis ordering failed: {e}")
            # Ultimate fallback - return shapes in original order
            return list(group_shape.shapes)

    def _get_shape_z_order(self, shape):
        """Extract z-order from shape XML."""
        try:
            if hasattr(shape, '_element') and hasattr(shape._element, 'xml'):
                xml_str = shape._element.xml

                # Look for z-order information in XML
                z_order_match = re.search(r'z-?order["\s]*[:=]["\s]*(\d+)', xml_str, re.IGNORECASE)
                if z_order_match:
                    return int(z_order_match.group(1))

                # Use shape ID as proxy
                id_match = re.search(r'id["\s]*=["\s]*["\'](\d+)["\']', xml_str)
                if id_match:
                    return int(id_match.group(1))

        except Exception:
            pass

        return 0

    def _extract_child_shape_info(self, shape_elem):
        """Extract information about a child shape from XML."""
        child_info = {
            'id': None,
            'name': None,
            'type': shape_elem.tag.split('}')[-1] if '}' in shape_elem.tag else shape_elem.tag,
            'z_order': 0
        }

        # Extract ID and name
        nv_props = (shape_elem.find('.//p:nvSpPr', self.namespaces) or
                    shape_elem.find('.//p:nvPicPr', self.namespaces) or
                    shape_elem.find('.//p:nvGraphicFramePr', self.namespaces) or
                    shape_elem.find('.//p:nvCxnSpPr', self.namespaces))

        if nv_props is not None:
            cnv_pr = nv_props.find('.//p:cNvPr', self.namespaces)
            if cnv_pr is not None:
                child_info['id'] = cnv_pr.get('id')
                child_info['name'] = cnv_pr.get('name', '')

        return child_info

    def _map_xml_children_to_pptx_children(self, xml_children, pptx_children):
        """Map XML child shape info to python-pptx child shapes."""
        ordered_children = []
        used_children = set()

        # Create lookup for python-pptx child shapes
        child_lookup = {}
        for child in pptx_children:
            child_id = self._get_shape_id(child)
            child_name = self._get_shape_name(child)

            if child_id:
                child_lookup[f"id_{child_id}"] = child
            if child_name:
                child_lookup[f"name_{child_name}"] = child

        # Match XML order to python-pptx children
        for xml_child in xml_children:
            matched_child = None

            # Try ID matching
            if xml_child['id']:
                matched_child = child_lookup.get(f"id_{xml_child['id']}")

            # Try name matching
            if not matched_child and xml_child['name']:
                matched_child = child_lookup.get(f"name_{xml_child['name']}")

            # Add if found and not already used
            if matched_child and matched_child not in used_children:
                ordered_children.append(matched_child)
                used_children.add(matched_child)

        # Add any remaining children
        for child in pptx_children:
            if child not in used_children:
                ordered_children.append(child)

        return ordered_children

