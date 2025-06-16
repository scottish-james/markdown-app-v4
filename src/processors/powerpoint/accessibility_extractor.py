"""
Simplified Accessibility Order Extractor with XML-first approach
XML-first approach with MarkItDown fallback when XML unavailable

ARCHITECTURE OVERVIEW:
This component determines the proper reading order of shapes on PowerPoint slides.
The key insight is using PowerPoint's internal XML structure rather than relying
on visual positioning, which provides more accurate accessibility ordering.

PROCESSING STRATEGIES:
1. XML-based semantic ordering: Uses PowerPoint's internal XML document order
2. Positional ordering: Fallback using top-left to bottom-right positioning
3. MarkItDown compatibility: Simple shape order when XML unavailable

XML PARSING APPROACH:
- Leverages PowerPoint's internal XML structure via python-pptx ._element.xml
- Parses document tree order for shapes within slide XML
- Extracts semantic roles from placeholder types and XML attributes
- Handles grouped shapes by recursively processing group contents

ERROR HANDLING STRATEGY:
- Defensive programming with try/catch around all XML access
- Graceful degradation: XML failure → positional order → simple shape order
- Logging of extraction method used for debugging
- No crashes on XML parsing failures

PERFORMANCE CONSIDERATIONS:
- XML parsing is more expensive than positional sorting
- Caches extraction method for debugging/monitoring
- Minimal XML processing - only extracts necessary attributes
- Early exits for empty/problematic slides
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET
import re


class AccessibilityOrderExtractor:
    """
    Extracts reading order from PowerPoint slides using XML-first approach.

    COMPONENT RESPONSIBILITIES:
    - Determine optimal reading order for slide shapes
    - Handle grouped shapes with internal ordering
    - Provide semantic role detection from XML
    - Fall back gracefully when XML unavailable

    PROCESSING MODES:
    - Semantic accessibility order: XML document order + semantic prioritization
    - Positional order: Top-to-bottom, left-to-right positioning
    - Simple order: Direct shape enumeration (MarkItDown compatibility)

    XML DEPENDENCIES:
    - Requires access to slide._element.xml or slide.element.xml
    - Uses ElementTree for XML parsing with PowerPoint namespaces
    - Handles missing XML gracefully with fallback strategies
    """

    def __init__(self, use_accessibility_order=True):
        """
        Initialize the accessibility extractor.

        CONFIGURATION:
        - use_accessibility_order: Controls semantic vs positional ordering
        - Tracking: Records extraction method used for debugging

        Args:
            use_accessibility_order (bool): Whether to use accessibility order vs positional
        """
        self.use_accessibility_order = use_accessibility_order
        self.last_extraction_method = "not_extracted"

        # XML namespaces for PowerPoint OOXML processing
        # These are standard PowerPoint namespace URIs - do not modify
        self.namespaces = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

    def get_slide_reading_order(self, slide, slide_number):
        """
        Main entry point for reading order extraction.
        Implements strategy pattern for different extraction methods.

        ALGORITHM SELECTION:
        1. Check configuration preference (accessibility vs positional)
        2. Verify XML availability for sophisticated processing
        3. Execute appropriate strategy with fallback handling
        4. Track method used for debugging and monitoring

        ERROR HANDLING:
        - XML parsing failures fall back to simpler methods
        - Never crashes - always returns usable shape order
        - Logs extraction method for post-processing analysis

        Args:
            slide: python-pptx Slide object
            slide_number (int): Slide number for debugging/logging

        Returns:
            list: Ordered list of shapes in reading order
        """
        if not self.use_accessibility_order:
            # Simple positional method - fastest but least sophisticated
            ordered_shapes = self._get_positional_ordered_shapes(slide)
            self.last_extraction_method = "positional_order"
            return ordered_shapes

        # Check XML availability before attempting sophisticated processing
        if not self._has_xml_access(slide):
            # No XML - fall back to MarkItDown approach (simple shape order)
            self.last_extraction_method = "markitdown_fallback"
            return list(slide.shapes)

        try:
            # XML available - use sophisticated accessibility order extraction
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
        XML availability check with multiple access pattern attempts.

        RATIONALE: Different python-pptx versions expose XML differently
        - Some use ._element.xml, others use .element.xml
        - Some slides may have corrupted/missing XML
        - This method tries all known access patterns

        PERFORMANCE: Fast check that avoids expensive XML parsing

        Args:
            slide: python-pptx Slide object

        Returns:
            bool: True if XML is accessible and usable
        """
        try:
            # Attempt to access slide XML using known patterns
            slide_xml = self._get_slide_xml(slide)
            return slide_xml is not None and len(slide_xml) > 0
        except Exception:
            return False

    def _get_semantic_accessibility_order(self, slide):
        """
        Advanced semantic ordering using XML document structure and semantic roles.

        ALGORITHM:
        1. Extract shapes in XML document order (PowerPoint's internal order)
        2. Process grouped shapes by extracting individual children
        3. Classify shapes by semantic role (title, subtitle, content, other)
        4. Reorder by semantic priority: titles → subtitles → content → other

        XML DOCUMENT ORDER: PowerPoint stores shapes in creation/editing order
        within the XML, which often reflects intended reading flow better than
        visual positioning.

        SEMANTIC CLASSIFICATION: Uses PowerPoint's placeholder types and shape
        names to identify semantic roles, providing more meaningful ordering.

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes ordered by semantic importance and document flow
        """
        # Step 1: Get all shapes in XML document order
        xml_ordered_shapes = self._get_xml_document_order(slide)

        # Step 2: Process groups by extracting children individually
        final_ordered_shapes = []
        for shape in xml_ordered_shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                # Groups: Extract internal reading order and add children
                group_children = self.get_reading_order_of_grouped_by_shape(shape)
                final_ordered_shapes.extend(group_children)
            else:
                # Regular shapes: Add directly
                final_ordered_shapes.append(shape)

        # Step 3: Separate by semantic importance for final ordering
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

        # Step 4: Return in semantic priority order
        return title_shapes + subtitle_shapes + content_shapes + other_shapes

    def _get_semantic_role_from_xml(self, shape):
        """
        Semantic role detection using PowerPoint's internal metadata.

        DETECTION HIERARCHY:
        1. PowerPoint placeholder types (most reliable)
        2. Shape names from XML attributes
        3. Fallback classification based on content type

        PLACEHOLDER TYPES: PowerPoint assigns semantic meaning to placeholder
        shapes (title, subtitle, body, etc.). This is the most reliable source
        of semantic information.

        SHAPE NAMES: User-assigned or PowerPoint-generated names can indicate
        semantic purpose (e.g., "Title 1", "Content Placeholder 2").

        CONTENT ANALYSIS: Final fallback based on whether shape contains text
        or other content types.

        Args:
            shape: python-pptx Shape object

        Returns:
            str: Semantic role ('title', 'subtitle', 'content', 'other')
        """
        # Priority 1: Check PowerPoint placeholder types (most reliable)
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
            # Placeholder access can fail - continue to fallback methods
            pass

        # Priority 2: Check shape names from XML
        try:
            if hasattr(shape, 'name') and shape.name:
                name_lower = shape.name.lower()

                if any(title_word in name_lower for title_word in ['title', 'heading', 'header']):
                    if 'subtitle' not in name_lower:
                        return "title"

                if 'subtitle' in name_lower:
                    return "subtitle"
        except Exception:
            # Shape name access can fail
            pass

        # Priority 3: Content-based classification (fallback)
        if hasattr(shape, 'text_frame') or hasattr(shape, 'text'):
            return "content"
        else:
            return "other"

    def _get_xml_document_order(self, slide):
        """
        Extract shapes in XML document order using ElementTree parsing.

        XML STRUCTURE: PowerPoint slides have a shape tree (spTree) that contains
        all shapes in document creation/editing order. This often reflects the
        intended reading flow better than visual positioning.

        PARSING STRATEGY:
        1. Extract raw XML from slide object
        2. Parse with ElementTree using PowerPoint namespaces
        3. Find shape tree (spTree) element
        4. Extract shape information in document order
        5. Map XML shape data back to python-pptx shape objects

        ERROR HANDLING: XML parsing can fail due to malformed XML, namespace
        issues, or missing elements. All failures are caught and re-raised
        with context for upstream handling.

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes in XML document order
        """
        try:
            # Step 1: Get the slide's raw XML
            slide_xml = self._get_slide_xml(slide)

            # Step 2: Parse XML to get shapes in document order
            xml_shape_info = self._parse_slide_xml_for_document_order(slide_xml)

            # Step 3: Map XML order to python-pptx shapes
            ordered_shapes = self._map_xml_to_pptx_shapes(xml_shape_info, slide.shapes)

            return ordered_shapes
        except Exception as e:
            raise Exception(f"XML document order extraction failed: {e}")

    def _get_positional_ordered_shapes(self, slide):
        """
        Simple fallback: positional ordering (top-to-bottom, left-to-right).

        ALGORITHM: Sort shapes by top position, then by left position for ties.
        This provides a reasonable reading order when XML analysis fails.

        COORDINATE SYSTEM: PowerPoint uses EMU (English Metric Units) for
        positioning. Smaller values = higher/further left.

        ERROR HANDLING: If shape positioning data is unavailable, defaults
        to (0,0) to avoid sort failures.

        PERFORMANCE: Much faster than XML parsing, useful for large presentations
        where speed is more important than perfect reading order.

        Args:
            slide: python-pptx Slide object

        Returns:
            list: Shapes ordered by position (top-to-bottom, left-to-right)
        """
        positioned_shapes = []
        for shape in slide.shapes:
            # Extract position with fallback for missing data
            if hasattr(shape, 'top') and hasattr(shape, 'left'):
                positioned_shapes.append((shape.top, shape.left, shape))
            else:
                # Default position to avoid sort errors
                positioned_shapes.append((0, 0, shape))

        # Sort by top position first, then left position
        positioned_shapes.sort(key=lambda x: (x[0], x[1]))
        return [shape for _, _, shape in positioned_shapes]

    def _parse_slide_xml_for_document_order(self, slide_xml):
        """
        Parse slide XML to extract shape information in document order.

        XML STRUCTURE ANALYSIS:
        - Root element contains slide content
        - spTree (shape tree) contains all shapes
        - Shape elements (sp, pic, graphicFrame, etc.) are in document order
        - Each shape has identifying information (ID, name, type)

        NAMESPACE HANDLING: Uses predefined PowerPoint namespaces for reliable
        element selection. Namespace prefixes must match PowerPoint OOXML spec.

        SHAPE TYPE DETECTION: Different XML elements represent different shape
        types (sp=shape, pic=picture, graphicFrame=table/chart, etc.).

        Args:
            slide_xml (str): Raw XML content of slide

        Returns:
            list: Shape information in document order with identification data
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

            # Include all PowerPoint shape types
            if tag_name in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent']:
                shape_info = self._extract_shape_info_from_xml(elem, idx)
                if shape_info:
                    shape_order_info.append(shape_info)

        return shape_order_info

    def _extract_shape_info_from_xml(self, shape_elem, order_index):
        """
        Extract identifying information from XML shape element.

        IDENTIFICATION STRATEGY:
        - ID: Unique numeric identifier assigned by PowerPoint
        - Name: User-visible name (can be user-assigned or auto-generated)
        - Type: XML element type (sp, pic, graphicFrame, etc.)
        - Text content: Preview for debugging/verification

        XML PATH PATTERNS:
        - Non-visual properties: nvSpPr, nvPicPr, nvGraphicFramePr, etc.
        - Common properties: cNvPr (contains ID and name)
        - Text content: a:t elements (drawing text)

        ERROR HANDLING: Missing elements don't cause failures - just leave
        fields as None/empty. This ensures processing continues even with
        malformed or incomplete XML.

        Args:
            shape_elem: XML element for shape
            order_index (int): Position in document order

        Returns:
            dict: Shape identification information for mapping
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
        # Different shape types have different non-visual property containers
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

        # Extract text content for verification/debugging
        text_elements = shape_elem.findall('.//a:t', self.namespaces)
        if text_elements:
            all_text = ' '.join([t.text for t in text_elements if t.text])
            if all_text.strip():
                shape_info['has_text'] = True
                shape_info['text_content'] = all_text.strip()[:50]  # First 50 chars

        return shape_info

    def _map_xml_to_pptx_shapes(self, xml_shape_info, pptx_shapes):
        """
        Map XML shape information back to python-pptx shape objects.

        MAPPING STRATEGY:
        1. Create lookup tables by ID and name
        2. Match XML shapes to python-pptx shapes using IDs (most reliable)
        3. Fall back to name matching for shapes without IDs
        4. Add any unmatched shapes at the end

        CHALLENGES:
        - XML and python-pptx may represent shapes differently
        - IDs are most reliable but not always present
        - Names can be duplicate or missing
        - Some shapes may exist in XML but not python-pptx (or vice versa)

        LOOKUP OPTIMIZATION: Uses dictionary lookup for O(1) matching instead
        of nested loops for better performance with large slide sets.

        Args:
            xml_shape_info (list): Shape info from XML parsing
            pptx_shapes: python-pptx shapes collection

        Returns:
            list: Ordered python-pptx shapes matching XML order
        """
        ordered_shapes = []
        used_shapes = set()

        # Create efficient lookup tables for matching
        shape_lookup = {}
        for shape in pptx_shapes:
            shape_id = self._get_shape_id(shape)
            shape_name = self._get_shape_name(shape)

            # Use prefixed keys to avoid ID/name collisions
            if shape_id:
                shape_lookup[f"id_{shape_id}"] = shape
            if shape_name:
                shape_lookup[f"name_{shape_name}"] = shape

        # Match XML order to python-pptx shapes
        for xml_info in xml_shape_info:
            matched_shape = None

            # Priority 1: Try ID matching (most reliable)
            if xml_info['id']:
                matched_shape = shape_lookup.get(f"id_{xml_info['id']}")

            # Priority 2: Try name matching (less reliable)
            if not matched_shape and xml_info['name']:
                matched_shape = shape_lookup.get(f"name_{xml_info['name']}")

            # Add if found and not already used
            if matched_shape and matched_shape not in used_shapes:
                ordered_shapes.append(matched_shape)
                used_shapes.add(matched_shape)

        # Add any remaining unmatched shapes at the end
        for shape in pptx_shapes:
            if shape not in used_shapes:
                ordered_shapes.append(shape)

        return ordered_shapes

    def _get_slide_xml(self, slide):
        """
        Extract raw XML from slide with multiple access pattern attempts.

        ACCESS PATTERNS: Different python-pptx versions expose XML differently:
        - slide._element.xml (common pattern)
        - slide.element.xml (alternative pattern)
        - slide.part._element.xml (deep access pattern)

        ERROR HANDLING: Tries all known patterns before giving up. Returns None
        on complete failure rather than crashing.

        Args:
            slide: python-pptx Slide object

        Returns:
            str|None: Raw XML content or None if inaccessible
        """
        try:
            # Pattern 1: Direct element access
            if hasattr(slide, '_element') and hasattr(slide._element, 'xml'):
                return slide._element.xml
            # Pattern 2: Public element access
            elif hasattr(slide, 'element') and hasattr(slide.element, 'xml'):
                return slide.element.xml
            else:
                # Pattern 3: Deep part access
                slide_part = slide.part if hasattr(slide, 'part') else None
                if slide_part and hasattr(slide_part, '_element'):
                    return slide_part._element.xml
                else:
                    raise Exception("Cannot access slide XML")
        except Exception:
            return None

    def _get_shape_id(self, shape):
        """
        Extract PowerPoint's internal shape ID from python-pptx shape.

        ID EXTRACTION: Shape IDs are buried in XML attributes. Uses regex
        to extract from XML string rather than parsing full XML tree for
        performance.

        REGEX PATTERN: Looks for cNvPr element with id attribute in XML.
        This is PowerPoint's standard pattern for shape identification.

        Args:
            shape: python-pptx Shape object

        Returns:
            str|None: Shape ID or None if not found
        """
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
        """
        Extract shape name with error handling.

        SIMPLE EXTRACTION: Shape names are directly accessible via python-pptx
        API, but access can still fail due to corrupted shape data.

        Args:
            shape: python-pptx Shape object

        Returns:
            str|None: Shape name or None if not accessible
        """
        try:
            if hasattr(shape, 'name') and shape.name:
                return shape.name
        except:
            pass
        return None

    def get_last_extraction_method(self):
        """
        Get the method used in the last extraction for debugging/monitoring.

        TRACKING: Records which extraction strategy was actually used:
        - semantic_accessibility_order: Full XML-based processing
        - positional_order: Fallback positioning
        - markitdown_fallback: Simple shape enumeration
        - xml_error_fallback: XML failed, using simple order

        Returns:
            str: Extraction method identifier
        """
        return self.last_extraction_method

    def get_reading_order_of_grouped_by_shape(self, group_shape):
        """
        Extract reading order of shapes within a group using XML or z-axis order.

        GROUP PROCESSING STRATEGY:
        1. Try XML-based group reading order (most accurate)
        2. Fall back to z-axis (stacking order) if XML fails
        3. Ultimate fallback to original shape order

        WHY GROUPS NEED SPECIAL HANDLING:
        Groups contain child shapes that may have their own internal ordering
        that's different from the parent slide's order. This is common in
        complex diagrams and grouped content.

        Args:
            group_shape: python-pptx GroupShape object

        Returns:
            list: Child shapes in proper reading order
        """
        try:
            # Primary strategy: XML-based group reading order
            xml_ordered_children = self._get_group_xml_reading_order(group_shape)
            if xml_ordered_children:
                return xml_ordered_children

        except Exception as e:
            print(f"XML group reading order failed: {e}")

        # Fallback strategy: Use z-axis (stacking order)
        return self._get_group_z_axis_order(group_shape)

    def _get_group_xml_reading_order(self, group_shape):
        """
        Extract child shapes from group XML in document order.

        GROUP XML STRUCTURE: Groups have their own internal XML structure
        with child shape elements. This preserves the creation/editing order
        of shapes within the group.

        PARSING APPROACH: Similar to slide-level XML parsing but focused
        on group-specific child elements.

        Args:
            group_shape: python-pptx GroupShape object

        Returns:
            list|None: Child shapes in XML document order or None if failed
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
        Get child shapes ordered by z-axis (stacking order) as fallback.

        Z-AXIS ORDERING: PowerPoint maintains stacking order (front-to-back)
        for shapes. This can be used as a proxy for reading order when XML
        parsing fails.

        EXTRACTION CHALLENGE: Z-order information is not directly exposed
        by python-pptx API, so we extract it from XML or use shape IDs as
        a proxy.

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
        """
        Extract z-order (stacking order) from shape XML with fallbacks.

        Z-ORDER EXTRACTION STRATEGY:
        1. Look for explicit z-order attributes in XML
        2. Use shape ID as proxy (IDs often correlate with creation order)
        3. Default to 0 if no order information available

        REGEX PATTERNS: Searches for various z-order representations in XML
        as PowerPoint may store this information in different formats.

        Args:
            shape: python-pptx Shape object

        Returns:
            int: Z-order value (higher = more forward)
        """
        try:
            if hasattr(shape, '_element') and hasattr(shape._element, 'xml'):
                xml_str = shape._element.xml

                # Look for explicit z-order information in XML
                z_order_match = re.search(r'z-?order["\s]*[:=]["\s]*(\d+)', xml_str, re.IGNORECASE)
                if z_order_match:
                    return int(z_order_match.group(1))

                # Fallback: Use shape ID as proxy for order
                id_match = re.search(r'id["\s]*=["\s]*["\'](\d+)["\']', xml_str)
                if id_match:
                    return int(id_match.group(1))

        except Exception:
            pass

        return 0  # Default z-order

    def _extract_child_shape_info(self, shape_elem):
        """
        Extract information about a child shape from group XML.

        CHILD SHAPE PROCESSING: Similar to slide-level shape extraction
        but focused on group context. Extracts identification info needed
        for mapping back to python-pptx objects.

        Args:
            shape_elem: XML element for child shape

        Returns:
            dict: Child shape identification information
        """
        child_info = {
            'id': None,
            'name': None,
            'type': shape_elem.tag.split('}')[-1] if '}' in shape_elem.tag else shape_elem.tag,
            'z_order': 0
        }

        # Extract ID and name using same pattern as parent slides
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
        """
        Map XML child shape info to python-pptx child shapes.

        CHILD MAPPING: Same strategy as parent slide mapping but applied
        to group children. Uses ID and name matching with fallbacks.

        Args:
            xml_children: List of child shape info from XML
            pptx_children: python-pptx group child shapes

        Returns:
            list: Ordered child shapes matching XML order
        """
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

            # Try ID matching first
            if xml_child['id']:
                matched_child = child_lookup.get(f"id_{xml_child['id']}")

            # Try name matching as fallback
            if not matched_child and xml_child['name']:
                matched_child = child_lookup.get(f"name_{xml_child['name']}")

            # Add if found and not already used
            if matched_child and matched_child not in used_children:
                ordered_children.append(matched_child)
                used_children.add(matched_child)

        # Add any remaining children that weren't matched
        for child in pptx_children:
            if child not in used_children:
                ordered_children.append(child)

        return ordered_children