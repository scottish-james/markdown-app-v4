def add_diagram_candidate_markers(self, markdown_content, structured_data):
    """
    Secondary process: Add diagram candidate markers based on lines and arrows
    """
    # Analyze each slide for diagram indicators
    diagram_slides = []

    for slide_idx, slide in enumerate(structured_data["slides"]):
        diagram_analysis = self.analyze_slide_for_diagram_shapes(slide)
        if diagram_analysis["is_likely_diagram"]:
            diagram_slides.append({
                "slide": slide_idx + 1,
                "analysis": diagram_analysis
            })

    # Add markers for slides with diagram indicators
    if diagram_slides:
        lines = markdown_content.split('\n')

        for diagram_slide in diagram_slides:
            slide_num = diagram_slide["slide"]
            analysis = diagram_slide["analysis"]

            # Create informative comment
            indicators = []
            if analysis["lines"]: indicators.append(f"lines: {analysis['lines']}")
            if analysis["arrows"]: indicators.append(f"arrows: {analysis['arrows']}")
            if analysis["connectors"]: indicators.append(f"connectors: {analysis['connectors']}")

            comment = f"\n<!-- DIAGRAM_DETECTED: {', '.join(indicators)} -->\n"

            # Find the slide and add comment after it
            slide_marker = f"<!-- Slide {slide_num} -->"
            for i, line in enumerate(lines):
                if slide_marker in line:
                    # Find the end of this slide's content
                    next_slide_idx = len(lines)
                    for j in range(i + 1, len(lines)):
                        if lines[j].strip().startswith('<!-- Slide '):
                            next_slide_idx = j
                            break

                    # Insert before next slide
                    lines.insert(next_slide_idx, comment)
                    break

        markdown_content = '\n'.join(lines)

    return markdown_content


def analyze_slide_for_diagram_shapes(self, slide_data):
    """
    Analyze a slide for diagram indicators: lines, arrows, connectors
    """
    analysis = {
        "is_likely_diagram": False,
        "lines": 0,
        "arrows": 0,
        "connectors": 0,
        "total_shapes": 0,
        "diagram_shapes": 0
    }

    # Check all content blocks on the slide
    for block in slide_data.get("content_blocks", []):
        self._analyze_block_for_diagram_shapes(block, analysis)

    # Determine if this looks like a diagram
    diagram_shape_count = analysis["lines"] + analysis["arrows"] + analysis["connectors"]
    analysis["diagram_shapes"] = diagram_shape_count

    # If we have lines/arrows/connectors, it's likely a diagram
    analysis["is_likely_diagram"] = diagram_shape_count > 0

    return analysis


def _analyze_block_for_diagram_shapes(self, block, analysis):
    """
    Recursively analyze a block and its shapes for diagram indicators
    """
    analysis["total_shapes"] += 1

    # Check if this block represents a line/arrow/connector
    # Note: We'd need to modify extract_shape_content to capture shape_type info
    shape_type = block.get("shape_type")
    auto_shape_type = block.get("auto_shape_type")

    if shape_type == "LINE":
        analysis["lines"] += 1
    elif shape_type == "CONNECTOR":
        """
PowerPoint Processor - Fixed and Complete
Maintains all original functionality while fixing bullet detection and adding slide title conversion
"""


from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
import json
import re
import os
from datetime import datetime
import xml.etree.ElementTree as ET


class PowerPointProcessor:
    """Complete PowerPoint processing with fixed bullet detection and slide title conversion"""

    def __init__(self):
        self.supported_formats = ['.pptx', '.ppt']

    def convert_pptx_to_markdown_enhanced(self, file_path, convert_slide_titles=True):
        """
        Main entry point: v14 text extraction + v19 diagram detection appended at end

        Args:
            file_path (str): Path to the PowerPoint file
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings
        """
        try:
            prs = Presentation(file_path)

            # Extract PowerPoint metadata first
            pptx_metadata = self.extract_pptx_metadata(prs, file_path)

            # Extract structured data (v14 approach)
            structured_data = self.extract_presentation_data(prs)

            # Convert to basic markdown (v14 approach)
            markdown = self.convert_structured_data_to_markdown(structured_data, convert_slide_titles)

            # Add PowerPoint metadata as comments for Claude to use
            markdown_with_metadata = self.add_pptx_metadata_for_claude(markdown, pptx_metadata)

            # APPEND v19 diagram analysis at the end
            diagram_analysis = self.analyze_structured_data_for_diagrams(structured_data)
            if diagram_analysis:
                markdown_with_metadata += "\n\n" + diagram_analysis

            return markdown_with_metadata
        except Exception as e:
            raise Exception(f"Error processing PowerPoint file: {str(e)}")

    def analyze_structured_data_for_diagrams(self, structured_data):
        """
        v19 diagram analysis system - analyze extracted structured data
        """
        try:
            diagram_slides = []

            for slide_idx, slide in enumerate(structured_data["slides"]):
                score_analysis = self.score_slide_for_diagram(slide)
                if score_analysis["probability"] >= 40:  # 40%+ probability threshold
                    diagram_slides.append({
                        "slide": slide_idx + 1,
                        "analysis": score_analysis
                    })

            # Generate detailed summary
            if diagram_slides:
                summary = "## DIAGRAM ANALYSIS (v19 Scoring System)\n\n"
                summary += "**Slides with potential diagrams:**\n\n"

                for slide_info in diagram_slides:
                    analysis = slide_info["analysis"]
                    summary += f"- **Slide {slide_info['slide']}**: {analysis['probability']}% probability "
                    summary += f"(Score: {analysis['total_score']}) - {', '.join(analysis['reasons'])}\n"
                    summary += f"  - Shapes: {analysis['shape_count']}, Lines: {analysis['line_count']}, Arrows: {analysis['arrow_count']}\n\n"

                return summary

            return None

        except Exception as e:
            return f"\n\n<!-- v19 Diagram analysis error: {e} -->"

    def score_slide_for_diagram(self, slide_data):
        """
        v19 scoring system: Score a slide for diagram probability using sophisticated rules
        """
        content_blocks = slide_data.get("content_blocks", [])

        # Collect all shapes and lines from structured data
        shapes = []
        lines = []
        arrows = []
        text_blocks = []

        for block in content_blocks:
            if block.get("type") == "line":
                lines.append(block)
            elif block.get("type") == "arrow":
                arrows.append(block)
            elif block.get("type") == "text":
                text_blocks.append(block)
                shapes.append(block)
            elif block.get("type") in ["shape", "image", "chart"]:
                shapes.append(block)
            elif block.get("type") == "group":
                # Recursively analyze group contents
                group_analysis = self._analyze_group_contents(block)
                shapes.extend(group_analysis["shapes"])
                lines.extend(group_analysis["lines"])
                arrows.extend(group_analysis["arrows"])
                text_blocks.extend(group_analysis["text_blocks"])

        # Calculate score based on v19 rules
        score = 0
        reasons = []

        # Rule 1: Line/Arrow threshold (20+ points each)
        if len(arrows) > 0:
            score += 20
            reasons.append(f"block_arrows:{len(arrows)}")

        if len(lines) >= 3:
            score += 20
            reasons.append(f"connector_lines:{len(lines)}")

        # Rule 2: Line-to-shape ratio (15 points)
        total_lines = len(lines) + len(arrows)
        if len(shapes) > 0:
            line_ratio = total_lines / len(shapes)
            if line_ratio >= 0.5:
                score += 15
                reasons.append(f"line_ratio:{line_ratio:.1f}")

        # Rule 3: Spatial layout analysis (10-15 points)
        layout_score = self._analyze_spatial_layout(shapes)
        score += layout_score["score"]
        if layout_score["score"] > 0:
            reasons.append(f"layout:{layout_score['type']}")

        # Rule 4: Shape variety (10-15 points)
        variety_score = self._analyze_shape_variety(shapes)
        score += variety_score
        if variety_score > 0:
            reasons.append(f"variety:{variety_score}")

        # Rule 5: Text density analysis (10 points)
        text_score = self._analyze_text_density(text_blocks)
        score += text_score
        if text_score > 0:
            reasons.append(f"short_text:{text_score}")

        # Rule 6: Flow patterns (20 points)
        flow_score = self._analyze_flow_patterns(shapes, lines, arrows, text_blocks)
        score += flow_score
        if flow_score > 0:
            reasons.append(f"flow_pattern:{flow_score}")

        # Negative indicators
        negative_score = self._analyze_negative_indicators(text_blocks, shapes)
        score += negative_score  # negative_score will be negative or 0
        if negative_score < 0:
            reasons.append(f"negatives:{negative_score}")

        # Convert score to probability
        if score >= 60:
            probability = 95
        elif score >= 40:
            probability = 75
        elif score >= 20:
            probability = 40
        else:
            probability = 10

        return {
            "total_score": score,
            "probability": probability,
            "reasons": reasons,
            "shape_count": len(shapes),
            "line_count": len(lines),
            "arrow_count": len(arrows)
        }

    def _analyze_group_contents(self, group_block):
        """Recursively analyze group contents for diagram elements"""
        result = {"shapes": [], "lines": [], "arrows": [], "text_blocks": []}

        for extracted_block in group_block.get("extracted_blocks", []):
            if extracted_block.get("type") == "line":
                result["lines"].append(extracted_block)
            elif extracted_block.get("type") == "arrow":
                result["arrows"].append(extracted_block)
            elif extracted_block.get("type") == "text":
                result["text_blocks"].append(extracted_block)
                result["shapes"].append(extracted_block)
            elif extracted_block.get("type") in ["shape", "image", "chart"]:
                result["shapes"].append(extracted_block)

        return result

    def _analyze_spatial_layout(self, shapes):
        """Analyze spatial layout patterns"""
        if len(shapes) < 3:
            return {"score": 0, "type": "insufficient"}

        positions = []
        for shape in shapes:
            pos = shape.get("position")
            if pos:
                positions.append((pos["top"], pos["left"]))

        if len(positions) < 3:
            return {"score": 0, "type": "no_position_data"}

        # Calculate spread
        tops = [p[0] for p in positions]
        lefts = [p[1] for p in positions]

        top_range = max(tops) - min(tops) if tops else 0
        left_range = max(lefts) - min(lefts) if lefts else 0

        # Check for grid-like arrangement
        unique_tops = len(set(round(t / 100000) for t in tops))  # Group by approximate position
        unique_lefts = len(set(round(l / 100000) for l in lefts))

        if unique_tops >= 2 and unique_lefts >= 2:
            return {"score": 15, "type": "grid_layout"}
        elif top_range > 1000000 and left_range > 1000000:
            return {"score": 10, "type": "spread_layout"}
        else:
            return {"score": 0, "type": "linear_layout"}

    def _analyze_shape_variety(self, shapes):
        """Analyze variety in shape types and sizes"""
        if len(shapes) < 2:
            return 0

        shape_types = set()
        sizes = []

        for shape in shapes:
            shape_types.add(shape.get("type", "unknown"))
            pos = shape.get("position")
            if pos:
                size = pos["width"] * pos["height"]
                sizes.append(size)

        score = 0

        # Multiple shape types
        if len(shape_types) >= 3:
            score += 15
        elif len(shape_types) >= 2:
            score += 10

        # Consistent sizing (indicates process flow)
        if len(sizes) >= 3:
            avg_size = sum(sizes) / len(sizes)
            variations = [abs(size - avg_size) / avg_size for size in sizes if avg_size > 0]
            if variations and max(variations) < 0.5:  # Less than 50% variation
                score += 5

        return score

    def _analyze_text_density(self, text_blocks):
        """Analyze text characteristics for diagram indicators"""
        if not text_blocks:
            return 0

        short_text_count = 0
        total_blocks = len(text_blocks)

        for block in text_blocks:
            # Count average words per paragraph
            total_words = 0
            para_count = 0

            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    words = len(clean_text.split())
                    total_words += words
                    para_count += 1

            if para_count > 0:
                avg_words = total_words / para_count
                if avg_words <= 5:  # Short labels
                    short_text_count += 1

        # Score based on percentage of short text blocks
        if total_blocks > 0:
            short_ratio = short_text_count / total_blocks
            if short_ratio >= 0.7:  # 70%+ short text
                return 10
            elif short_ratio >= 0.5:  # 50%+ short text
                return 5

        return 0

    def _analyze_flow_patterns(self, shapes, lines, arrows, text_blocks):
        """Analyze for flow patterns and process keywords"""
        score = 0

        # Check for start/end keywords
        flow_keywords = ["start", "begin", "end", "finish", "process", "step", "decision"]
        action_words = ["create", "update", "check", "verify", "send", "receive", "analyze"]

        all_text = ""
        for block in text_blocks:
            for para in block.get("paragraphs", []):
                all_text += " " + para.get("clean_text", "").lower()

        flow_matches = sum(1 for keyword in flow_keywords if keyword in all_text)
        action_matches = sum(1 for keyword in action_words if keyword in all_text)

        if flow_matches >= 2:
            score += 20
        elif flow_matches >= 1:
            score += 10

        if action_matches >= 3:
            score += 10

        # Bonus for having both shapes and connecting elements
        if len(shapes) >= 3 and (len(lines) > 0 or len(arrows) > 0):
            score += 15

        return score

    def _analyze_negative_indicators(self, text_blocks, shapes):
        """Check for negative indicators that suggest NOT a diagram"""
        score = 0

        # Check for long paragraphs
        long_text_count = 0
        bullet_count = 0

        for block in text_blocks:
            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    word_count = len(clean_text.split())
                    if word_count > 20:  # Long paragraph
                        long_text_count += 1

                    # Check for bullet points
                    if para.get("hints", {}).get("is_bullet", False):
                        bullet_count += 1

        # Penalize long text
        if long_text_count >= 2:
            score -= 15

        # Penalize if mostly bullet points
        total_paras = sum(len(block.get("paragraphs", [])) for block in text_blocks)
        if total_paras > 0 and bullet_count / total_paras > 0.8:
            score -= 10

        # Penalize single column layout (all shapes vertically aligned)
        if len(shapes) >= 3:
            positions = [s.get("position") for s in shapes if s.get("position")]
            if len(positions) >= 3:
                lefts = [p["left"] for p in positions]
                left_variance = max(lefts) - min(lefts) if lefts else 0
                if left_variance < 500000:  # Very narrow horizontal spread
                    score -= 10

        return score

    def is_arrow_shape(self, auto_shape_type):
        """Check if an auto shape type is an arrow"""
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

    def extract_pptx_metadata(self, presentation, file_path):
        """Extract comprehensive metadata from PowerPoint file"""
        metadata = {}

        try:
            # Core properties from PowerPoint
            core_props = presentation.core_properties

            # Basic file info
            metadata['filename'] = os.path.basename(file_path)
            metadata['file_size'] = os.path.getsize(file_path) if os.path.exists(file_path) else None

            # Document properties
            metadata['title'] = getattr(core_props, 'title', '') or ''
            metadata['author'] = getattr(core_props, 'author', '') or ''
            metadata['subject'] = getattr(core_props, 'subject', '') or ''
            metadata['keywords'] = getattr(core_props, 'keywords', '') or ''
            metadata['comments'] = getattr(core_props, 'comments', '') or ''
            metadata['category'] = getattr(core_props, 'category', '') or ''
            metadata['content_status'] = getattr(core_props, 'content_status', '') or ''
            metadata['language'] = getattr(core_props, 'language', '') or ''
            metadata['version'] = getattr(core_props, 'version', '') or ''

            # Dates
            metadata['created'] = getattr(core_props, 'created', None)
            metadata['modified'] = getattr(core_props, 'modified', None)
            metadata['last_modified_by'] = getattr(core_props, 'last_modified_by', '') or ''
            metadata['last_printed'] = getattr(core_props, 'last_printed', None)

            # Revision and identifier
            metadata['revision'] = getattr(core_props, 'revision', None)
            metadata['identifier'] = getattr(core_props, 'identifier', '') or ''

            # Presentation-specific info
            metadata['slide_count'] = len(presentation.slides)

            # Try to extract slide master themes/layouts
            try:
                slide_masters = presentation.slide_masters
                if slide_masters:
                    metadata['slide_master_count'] = len(slide_masters)
                    # Get layout names if available
                    layout_names = []
                    for master in slide_masters:
                        for layout in master.slide_layouts:
                            if hasattr(layout, 'name') and layout.name:
                                layout_names.append(layout.name)
                    metadata['layout_types'] = ', '.join(set(layout_names)) if layout_names else ''
            except:
                metadata['slide_master_count'] = 0
                metadata['layout_types'] = ''

            # Application that created the file
            try:
                app_props = presentation.app_properties if hasattr(presentation, 'app_properties') else None
                if app_props:
                    metadata['application'] = getattr(app_props, 'application', '') or ''
                    metadata['app_version'] = getattr(app_props, 'app_version', '') or ''
                    metadata['company'] = getattr(app_props, 'company', '') or ''
                    metadata['doc_security'] = getattr(app_props, 'doc_security', None)
            except:
                metadata['application'] = ''
                metadata['app_version'] = ''
                metadata['company'] = ''
                metadata['doc_security'] = None

        except Exception as e:
            print(f"Warning: Could not extract some metadata: {e}")

        return metadata

    def add_pptx_metadata_for_claude(self, markdown_content, metadata):
        """Add PowerPoint metadata as comments for Claude to incorporate"""
        # Format metadata for Claude
        metadata_comments = "\n<!-- POWERPOINT METADATA FOR CLAUDE:\n"

        if metadata.get('title'):
            metadata_comments += f"Document Title: {metadata['title']}\n"
        if metadata.get('author'):
            metadata_comments += f"Author: {metadata['author']}\n"
        if metadata.get('subject'):
            metadata_comments += f"Subject: {metadata['subject']}\n"
        if metadata.get('keywords'):
            metadata_comments += f"Keywords: {metadata['keywords']}\n"
        if metadata.get('category'):
            metadata_comments += f"Category: {metadata['category']}\n"
        if metadata.get('comments'):
            metadata_comments += f"Document Comments: {metadata['comments']}\n"
        if metadata.get('created'):
            metadata_comments += f"Created Date: {metadata['created']}\n"
        if metadata.get('modified'):
            metadata_comments += f"Last Modified: {metadata['modified']}\n"
        if metadata.get('last_modified_by'):
            metadata_comments += f"Last Modified By: {metadata['last_modified_by']}\n"
        if metadata.get('version'):
            metadata_comments += f"Version: {metadata['version']}\n"
        if metadata.get('application'):
            metadata_comments += f"Created With: {metadata['application']}\n"
        if metadata.get('company'):
            metadata_comments += f"Company: {metadata['company']}\n"
        if metadata.get('language'):
            metadata_comments += f"Language: {metadata['language']}\n"
        if metadata.get('content_status'):
            metadata_comments += f"Content Status: {metadata['content_status']}\n"

        # File info
        metadata_comments += f"Filename: {metadata.get('filename', 'unknown')}\n"
        metadata_comments += f"Slide Count: {metadata.get('slide_count', 0)}\n"
        if metadata.get('slide_size'):
            metadata_comments += f"Slide Size: {metadata['slide_size']}\n"
        if metadata.get('layout_types'):
            metadata_comments += f"Layout Types: {metadata['layout_types']}\n"

        metadata_comments += "-->\n"

        # Add metadata at the beginning
        return metadata_comments + markdown_content

    def extract_presentation_data(self, presentation):
        """Extract all content with minimal processing"""
        data = {
            "total_slides": len(presentation.slides),
            "slides": []
        }

        for slide_idx, slide in enumerate(presentation.slides, 1):
            slide_data = self.extract_slide_data(slide, slide_idx)
            data["slides"].append(slide_data)

        return data

    def extract_slide_data(self, slide, slide_number):
        """Extract slide content in reading order"""
        # Get shapes in reading order
        positioned_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, 'top') and hasattr(shape, 'left'):
                positioned_shapes.append((shape.top, shape.left, shape))
            else:
                positioned_shapes.append((0, 0, shape))

        positioned_shapes.sort(key=lambda x: (x[0], x[1]))

        slide_data = {
            "slide_number": slide_number,
            "content_blocks": []
        }

        for _, _, shape in positioned_shapes:
            block = self.extract_shape_content(shape)
            if block:
                slide_data["content_blocks"].append(block)

        return slide_data

    def extract_shape_content(self, shape):
        """Extract shape content with proper type detection - v14 approach BUT capture shape info for diagram analysis"""
        # Capture basic shape info for later diagram analysis (with safe error handling)
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

        # Safely get shape type
        try:
            if hasattr(shape, 'shape_type'):
                shape_info["shape_type"] = str(shape.shape_type).split('.')[-1]  # Get just the name part
        except:
            shape_info["shape_type"] = "unknown"

        # Check for auto shape type (for arrows and special shapes)
        try:
            if hasattr(shape, 'auto_shape_type'):
                shape_info["auto_shape_type"] = str(shape.auto_shape_type).split('.')[-1]
        except:
            pass

        # MAIN EXTRACTION - v14 approach that works
        content_block = None

        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                content_block = self.extract_image(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                content_block = self.extract_table(shape.table)
            elif hasattr(shape, 'has_chart') and shape.has_chart:
                content_block = self.extract_chart(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                content_block = self.extract_group(shape)
            elif hasattr(shape, 'text_frame') and shape.text_frame:
                content_block = self.extract_text_frame_fixed(shape.text_frame, shape)
            elif hasattr(shape, 'text') and shape.text:
                content_block = self.extract_plain_text(shape)
        except Exception as e:
            print(f"Warning: Error extracting shape content: {e}")
            return None

        # DIAGRAM ANALYSIS - add shape info for diagram detection (with safe checks)
        if not content_block:
            # For shapes without text content, create minimal blocks for diagram analysis
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.LINE:
                    content_block = {"type": "line", "line_type": "simple"}
                elif shape.shape_type == MSO_SHAPE_TYPE.CONNECTOR:
                    content_block = {"type": "line", "line_type": "connector"}
                elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
                    content_block = {"type": "line", "line_type": "freeform"}
                elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    if self.is_arrow_shape(shape_info["auto_shape_type"]):
                        content_block = {"type": "arrow", "arrow_type": shape_info["auto_shape_type"]}
                    else:
                        content_block = {"type": "shape", "shape_subtype": "auto_shape"}
                else:
                    content_block = {"type": "shape", "shape_subtype": "generic"}
            except Exception as e:
                # Fallback for any shape type issues
                content_block = {"type": "shape", "shape_subtype": "unknown"}

        # Add shape analysis info to content block for diagram detection
        if content_block:
            try:
                content_block.update(shape_info)
            except Exception as e:
                print(f"Warning: Error adding shape info: {e}")

        return content_block

    def is_arrow_shape(self, auto_shape_type):
        """Check if an auto shape type is an arrow"""
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

    def extract_text_frame_fixed(self, text_frame, shape):
        """Fixed text extraction with proper bullet detection"""
        if not text_frame.paragraphs:
            return None

        block = {
            "type": "text",
            "paragraphs": [],
            "shape_hyperlink": self.extract_shape_hyperlink(shape)
        }

        for para_idx, para in enumerate(text_frame.paragraphs):
            if not para.text.strip():
                continue

            para_data = self.process_paragraph_fixed(para)
            if para_data:
                block["paragraphs"].append(para_data)

        return block if block["paragraphs"] else None

    def process_paragraph_fixed(self, para):
        """Fixed paragraph processing with reliable bullet detection"""
        raw_text = para.text
        if not raw_text.strip():
            return None

        # First, check if PowerPoint knows this is a bullet
        ppt_level = getattr(para, 'level', None)

        # Check XML for bullet formatting
        is_ppt_bullet = False
        xml_level = None

        try:
            if hasattr(para, '_p') and para._p is not None:
                xml_str = str(para._p.xml)
                # Look for bullet indicators
                if any(indicator in xml_str for indicator in ['buChar', 'buAutoNum', 'buFont']):
                    is_ppt_bullet = True
                    # Try to extract level
                    import re
                    level_match = re.search(r'lvl="(\d+)"', xml_str)
                    if level_match:
                        xml_level = int(level_match.group(1))
        except:
            pass

        # Determine final bullet level
        bullet_level = -1
        if is_ppt_bullet:
            bullet_level = xml_level if xml_level is not None else (ppt_level if ppt_level is not None else 0)
        elif ppt_level is not None:
            # PowerPoint says it has a level, trust it
            bullet_level = ppt_level

        # Check for manual bullets and numbered lists
        clean_text = raw_text.strip()
        manual_bullet = self.is_manual_bullet(clean_text)
        numbered = self.is_numbered_list(clean_text)

        # If we found a manual bullet but PowerPoint didn't recognize it
        if manual_bullet and bullet_level < 0:
            # Estimate level from indentation
            leading_spaces = len(raw_text) - len(raw_text.lstrip())
            bullet_level = min(leading_spaces // 2, 6)
            clean_text = self.remove_bullet_char(clean_text)
        elif bullet_level >= 0:
            # Remove any manual bullet chars if PowerPoint formatted it
            clean_text = self.remove_bullet_char(clean_text)
        elif numbered:
            clean_text = self.remove_number_prefix(clean_text)

        # Extract formatted runs - THIS IS KEY!
        formatted_runs = self.extract_runs_with_text_preservation(para.runs, clean_text, bullet_level >= 0 or numbered)

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
                "likely_heading": self.is_likely_heading(clean_text)
            }
        }

        return para_data

    def extract_runs_with_text_preservation(self, runs, clean_text, has_prefix_removed):
        """Extract runs while preserving formatting after bullet/number removal"""
        if not runs:
            return [{"text": clean_text, "bold": False, "italic": False, "hyperlink": None}]

        formatted_runs = []

        # If we removed a prefix (bullet/number), we need to adjust the runs
        if has_prefix_removed:
            # Find where the clean text starts in the original runs
            full_text = "".join(run.text for run in runs)

            # Find the start position of clean_text in full_text
            # This is tricky because clean_text has had prefixes removed
            start_pos = -1
            for i in range(len(full_text)):
                remaining = full_text[i:].strip()
                if remaining == clean_text:
                    start_pos = i
                    break

            if start_pos == -1:
                # Fallback: just process runs normally
                start_pos = 0

            # Now process runs, skipping content before start_pos
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
                    formatted_runs.append(self.extract_run_formatting(run, run_text))

                char_count += len(run.text)
        else:
            # No prefix removed, process runs normally
            for run in runs:
                if run.text:
                    formatted_runs.append(self.extract_run_formatting(run, run.text))

        return formatted_runs

    def extract_run_formatting(self, run, text_override=None):
        """Extract formatting from a single run"""
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
                run_data["hyperlink"] = self.fix_url(run.hyperlink.address)
        except:
            pass

        return run_data

    def is_manual_bullet(self, text):
        """Check if text starts with a manual bullet character"""
        if not text:
            return False
        bullet_chars = '•◦▪▫‣·○■□→►✓✗-*+※◆◇'
        return text[0] in bullet_chars

    def is_numbered_list(self, text):
        """Check if text starts with a number pattern"""
        patterns = [
            r'^\d+[\.\)]\s+',  # 1. or 1)
            r'^[a-zA-Z][\.\)]\s+',  # a. or A)
            r'^[ivxlcdm]+[\.\)]\s+',  # Roman numerals (lowercase)
            r'^[IVXLCDM]+[\.\)]\s+',  # Roman numerals (uppercase)
        ]
        return any(re.match(pattern, text) for pattern in patterns)

    def remove_bullet_char(self, text):
        """Remove bullet characters from start of text"""
        if not text:
            return text
        # Remove common bullet chars and following spaces
        return re.sub(r'^[•◦▪▫‣·○■□→►✓✗\-\*\+※◆◇]\s*', '', text)

    def remove_number_prefix(self, text):
        """Remove number prefix from text"""
        return re.sub(r'^[^\s]+\s+', '', text)

    def is_likely_heading(self, text):
        """Determine if text is likely a heading"""
        if not text or len(text) > 150:
            return False

        # All caps
        if text.isupper() and len(text) > 2:
            return True

        # Short text without ending punctuation
        if len(text) < 80 and not text.endswith(('.', '!', '?', ';', ':', ',')):
            return True

        return False

    def extract_plain_text(self, shape):
        """Extract plain text from shape"""
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
            "shape_hyperlink": self.extract_shape_hyperlink(shape)
        }

    def _analyze_plain_text_hints(self, text):
        """Analyze plain text for formatting hints"""
        if not text:
            return {}

        stripped = text.strip()

        # Check each line for bullets
        lines = text.split('\n')
        has_bullets = False
        for line in lines:
            if line.strip() and self.is_manual_bullet(line.strip()):
                has_bullets = True
                break

        return {
            "has_powerpoint_level": False,
            "powerpoint_level": None,
            "bullet_level": -1,
            "is_bullet": has_bullets,
            "is_numbered": any(self.is_numbered_list(line.strip()) for line in lines if line.strip()),
            "starts_with_bullet": stripped and self.is_manual_bullet(stripped),
            "starts_with_number": bool(re.match(r'^\s*\d+[\.\)]\s', text)),
            "short_text": len(stripped) < 100,
            "all_caps": stripped.isupper() if stripped else False,
            "likely_heading": self.is_likely_heading(stripped)
        }

    def extract_image(self, shape):
        """Extract image info with proper alt text extraction"""
        alt_text = "Image"

        try:
            # Try multiple methods to get alt text
            if hasattr(shape, 'alt_text') and shape.alt_text:
                alt_text = shape.alt_text
            elif hasattr(shape, 'image') and hasattr(shape.image, 'alt_text') and shape.image.alt_text:
                alt_text = shape.image.alt_text
            elif hasattr(shape, '_element'):
                # Try to extract from XML
                try:
                    xml_str = str(shape._element.xml) if hasattr(shape._element, 'xml') else ""
                    if xml_str:
                        root = ET.fromstring(xml_str)
                        # Look for description attributes
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
            "hyperlink": self.extract_shape_hyperlink(shape)
        }

    def extract_table(self, table):
        """Extract table data"""
        if not table.rows:
            return None

        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                # Extract cell text with formatting
                cell_content = ""
                if hasattr(cell, 'text_frame') and cell.text_frame:
                    cell_paras = []
                    for para in cell.text_frame.paragraphs:
                        if para.text.strip():
                            # Process paragraph for bullets
                            para_processed = self.process_paragraph_fixed(para)
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
        """Extract chart/diagram information"""
        try:
            chart = shape.chart
            chart_data = {
                "type": "chart",
                "chart_type": str(chart.chart_type) if hasattr(chart, 'chart_type') else "unknown",
                "title": "",
                "data_points": [],
                "categories": [],
                "series": [],
                "hyperlink": self.extract_shape_hyperlink(shape)
            }

            # Try to get chart title
            try:
                if hasattr(chart, 'chart_title') and chart.chart_title and hasattr(chart.chart_title, 'text_frame'):
                    chart_data["title"] = chart.chart_title.text_frame.text.strip()
            except:
                pass

            # Try to extract data for potential Mermaid conversion
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
            # Fallback for charts we can't parse
            return {
                "type": "chart",
                "chart_type": "unknown",
                "title": "Chart",
                "data_points": [],
                "categories": [],
                "series": [],
                "hyperlink": self.extract_shape_hyperlink(shape)
            }

    def extract_group(self, shape):
        """Extract content from grouped shapes - EXACT v14 approach that was working"""
        try:
            # For grouped shapes, extract text from all child shapes
            extracted_blocks = []

            for child_shape in shape.shapes:
                # Extract text directly from each child shape - EXACTLY like v14
                if hasattr(child_shape, 'text_frame') and child_shape.text_frame:
                    text_block = self.extract_text_frame_fixed(child_shape.text_frame, child_shape)
                    if text_block:
                        extracted_blocks.append(text_block)
                elif hasattr(child_shape, 'text') and child_shape.text:
                    text_block = self.extract_plain_text(child_shape)
                    if text_block:
                        extracted_blocks.append(text_block)
                elif child_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image_block = self.extract_image(child_shape)
                    if image_block:
                        extracted_blocks.append(image_block)
                elif child_shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    table_block = self.extract_table(child_shape.table)
                    if table_block:
                        extracted_blocks.append(table_block)
                elif hasattr(child_shape, 'has_chart') and child_shape.has_chart:
                    chart_block = self.extract_chart(child_shape)
                    if chart_block:
                        extracted_blocks.append(chart_block)
                # Handle nested groups recursively
                elif child_shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    nested_group = self.extract_group(child_shape)
                    if nested_group and nested_group.get("extracted_blocks"):
                        extracted_blocks.extend(nested_group["extracted_blocks"])

            # Return a simplified group structure
            if extracted_blocks:
                return {
                    "type": "group",
                    "extracted_blocks": extracted_blocks,
                    "hyperlink": self.extract_shape_hyperlink(shape)
                }

            return None

        except Exception as e:
            print(f"Error extracting group: {e}")
            return None

    def analyze_diagram_pattern(self, shapes):
        """Analyze shapes to determine if they form a recognizable diagram pattern - be very conservative"""
        if not shapes:
            return "text_group"  # Changed default

        text_shapes = [s for s in shapes if s.get("type") == "text"]
        non_text_shapes = [s for s in shapes if s.get("type") != "text"]

        # Look for keywords to identify diagram type
        flowchart_keywords = ["start", "end", "process", "decision", "flow", "workflow", "step", "next", "previous"]
        org_keywords = ["manager", "director", "ceo", "team", "department", "reports to", "supervisor"]

        all_text = " ".join([
            " ".join([p.get("clean_text", "") for p in shape.get("paragraphs", [])])
            for shape in text_shapes
        ]).lower()

        # Only treat as diagram if we have very strong indicators
        strong_flowchart_match = sum(1 for keyword in flowchart_keywords if keyword in all_text) >= 2
        strong_org_match = sum(1 for keyword in org_keywords if keyword in all_text) >= 2
        has_non_text_shapes = len(non_text_shapes) > 0
        many_shapes = len(shapes) >= 6  # Raised threshold

        if strong_flowchart_match and (has_non_text_shapes or many_shapes):
            return "flowchart"
        elif strong_org_match and (has_non_text_shapes or many_shapes):
            return "org_chart"
        elif has_non_text_shapes and len(shapes) >= 4:
            return "diagram"

        # Default to text group for most cases
        return "text_group"

    def extract_shape_hyperlink(self, shape):
        """Extract shape-level hyperlink"""
        try:
            if hasattr(shape, 'click_action') and shape.click_action:
                if hasattr(shape.click_action, 'hyperlink') and shape.click_action.hyperlink:
                    if shape.click_action.hyperlink.address:
                        return self.fix_url(shape.click_action.hyperlink.address)
        except:
            pass
        return None

    def fix_url(self, url):
        """Fix URLs by adding schemes if missing"""
        if not url:
            return url

        if '@' in url and not url.startswith('mailto:'):
            return f"mailto:{url}"

        if not url.startswith(('http://', 'https://', 'mailto:', 'tel:', 'ftp://', '#')):
            if url.startswith('www.') or any(
                    domain in url.lower() for domain in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                return f"https://{url}"

        return url

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """Convert structured data to markdown"""
        markdown_parts = []

        for slide in data["slides"]:
            # Add slide marker
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            for block in slide["content_blocks"]:
                if block["type"] == "text":
                    markdown_parts.append(self.convert_text_block_to_markdown(block))
                elif block["type"] == "table":
                    markdown_parts.append(self.convert_table_to_markdown(block))
                elif block["type"] == "image":
                    markdown_parts.append(self.convert_image_to_markdown(block))
                elif block["type"] == "chart":
                    markdown_parts.append(self.convert_chart_to_markdown(block))
                elif block["type"] == "group":
                    markdown_parts.append(self.convert_group_to_markdown(block))

    def convert_structured_data_to_markdown(self, data, convert_slide_titles=True):
        """Convert structured data to markdown"""
        markdown_parts = []

        for slide in data["slides"]:
            # Add slide marker
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            for block in slide["content_blocks"]:
                if block["type"] == "text":
                    markdown_parts.append(self.convert_text_block_to_markdown(block))
                elif block["type"] == "table":
                    markdown_parts.append(self.convert_table_to_markdown(block))
                elif block["type"] == "image":
                    markdown_parts.append(self.convert_image_to_markdown(block))
                elif block["type"] == "chart":
                    markdown_parts.append(self.convert_chart_to_markdown(block))
                elif block["type"] == "group":
                    markdown_parts.append(self.convert_group_to_markdown(block))

        markdown_content = "\n\n".join(filter(None, markdown_parts))

        # Post-process to convert slide titles from bullets to H1 headings if requested
        if convert_slide_titles:
            markdown_content = self.convert_slide_titles_to_headings(markdown_content)

        return markdown_content

    def convert_slide_titles_to_headings(self, markdown_content):
        """
        Post-process markdown to convert slide titles from bullet points to H1 headings.

        This function identifies likely slide titles by looking for bullet points that appear
        immediately after slide markers and have title-like characteristics.
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
                    if self.is_likely_slide_title(next_line):
                        # Convert bullet to H1 heading
                        title_text = self.extract_title_from_bullet(next_line)
                        processed_lines.append(f"\n# {title_text}")
                        i = j  # Skip the original bullet line
                    else:
                        i = j - 1  # Process the next line normally
                else:
                    break

            i += 1

        return '\n'.join(processed_lines)

    def add_diagram_candidate_markers(self, markdown_content, structured_data):
        """
        Secondary process: Add diagram candidate markers based on scoring system
        """
        # Analyze each slide for diagram probability
        diagram_slides = []

        for slide_idx, slide in enumerate(structured_data["slides"]):
            score_analysis = self.score_slide_for_diagram(slide)
            if score_analysis["probability"] >= 40:  # 40%+ probability threshold
                diagram_slides.append({
                    "slide": slide_idx + 1,
                    "analysis": score_analysis
                })

        # Add markers for slides with high diagram probability
        if diagram_slides:
            lines = markdown_content.split('\n')

            for diagram_slide in diagram_slides:
                slide_num = diagram_slide["slide"]
                analysis = diagram_slide["analysis"]

                # Create detailed comment
                comment = f"\n<!-- DIAGRAM_DETECTED: probability={analysis['probability']}%, score={analysis['total_score']}, reasons={', '.join(analysis['reasons'])} -->\n"

                # Find the slide and add comment after it
                slide_marker = f"<!-- Slide {slide_num} -->"
                for i, line in enumerate(lines):
                    if slide_marker in line:
                        # Find the end of this slide's content
                        next_slide_idx = len(lines)
                        for j in range(i + 1, len(lines)):
                            if lines[j].strip().startswith('<!-- Slide '):
                                next_slide_idx = j
                                break

                        # Insert before next slide
                        lines.insert(next_slide_idx, comment)
                        break

            markdown_content = '\n'.join(lines)

        return markdown_content

    def score_slide_for_diagram(self, slide_data):
        """
        Score a slide for diagram probability using our rules
        """
        content_blocks = slide_data.get("content_blocks", [])

        # Collect all shapes and lines
        shapes = []
        lines = []
        arrows = []
        text_blocks = []

        for block in content_blocks:
            if block.get("type") == "line":
                lines.append(block)
            elif block.get("type") == "arrow":
                arrows.append(block)
            elif block.get("type") == "text":
                text_blocks.append(block)
                shapes.append(block)
            elif block.get("type") in ["shape", "image", "chart"]:
                shapes.append(block)
            elif block.get("type") == "group":
                # Recursively analyze group contents
                group_analysis = self._analyze_group_contents(block)
                shapes.extend(group_analysis["shapes"])
                lines.extend(group_analysis["lines"])
                arrows.extend(group_analysis["arrows"])
                text_blocks.extend(group_analysis["text_blocks"])

        # Calculate score based on rules
        score = 0
        reasons = []

        # Rule 1: Line/Arrow threshold (20+ points each)
        if len(arrows) > 0:
            score += 20
            reasons.append(f"block_arrows:{len(arrows)}")

        if len(lines) >= 3:
            score += 20
            reasons.append(f"connector_lines:{len(lines)}")

        # Rule 2: Line-to-shape ratio (15 points)
        total_lines = len(lines) + len(arrows)
        if len(shapes) > 0:
            line_ratio = total_lines / len(shapes)
            if line_ratio >= 0.5:
                score += 15
                reasons.append(f"line_ratio:{line_ratio:.1f}")

        # Rule 3: Spatial layout analysis (10-15 points)
        layout_score = self._analyze_spatial_layout(shapes)
        score += layout_score["score"]
        if layout_score["score"] > 0:
            reasons.append(f"layout:{layout_score['type']}")

        # Rule 4: Shape variety (10-15 points)
        variety_score = self._analyze_shape_variety(shapes)
        score += variety_score
        if variety_score > 0:
            reasons.append(f"variety:{variety_score}")

        # Rule 5: Text density analysis (10 points)
        text_score = self._analyze_text_density(text_blocks)
        score += text_score
        if text_score > 0:
            reasons.append(f"short_text:{text_score}")

        # Rule 6: Flow patterns (20 points)
        flow_score = self._analyze_flow_patterns(shapes, lines, arrows, text_blocks)
        score += flow_score
        if flow_score > 0:
            reasons.append(f"flow_pattern:{flow_score}")

        # Negative indicators
        negative_score = self._analyze_negative_indicators(text_blocks, shapes)
        score += negative_score  # negative_score will be negative or 0
        if negative_score < 0:
            reasons.append(f"negatives:{negative_score}")

        # Convert score to probability
        if score >= 60:
            probability = 95
        elif score >= 40:
            probability = 75
        elif score >= 20:
            probability = 40
        else:
            probability = 10

        return {
            "total_score": score,
            "probability": probability,
            "reasons": reasons,
            "shape_count": len(shapes),
            "line_count": len(lines),
            "arrow_count": len(arrows)
        }

    def _analyze_group_contents(self, group_block):
        """Recursively analyze group contents for diagram elements"""
        result = {"shapes": [], "lines": [], "arrows": [], "text_blocks": []}

        for extracted_block in group_block.get("extracted_blocks", []):
            if extracted_block.get("type") == "line":
                result["lines"].append(extracted_block)
            elif extracted_block.get("type") == "arrow":
                result["arrows"].append(extracted_block)
            elif extracted_block.get("type") == "text":
                result["text_blocks"].append(extracted_block)
                result["shapes"].append(extracted_block)
            elif extracted_block.get("type") in ["shape", "image", "chart"]:
                result["shapes"].append(extracted_block)

        return result

    def _analyze_spatial_layout(self, shapes):
        """Analyze spatial layout patterns"""
        if len(shapes) < 3:
            return {"score": 0, "type": "insufficient"}

        positions = []
        for shape in shapes:
            pos = shape.get("position")
            if pos:
                positions.append((pos["top"], pos["left"]))

        if len(positions) < 3:
            return {"score": 0, "type": "no_position_data"}

        # Calculate spread
        tops = [p[0] for p in positions]
        lefts = [p[1] for p in positions]

        top_range = max(tops) - min(tops) if tops else 0
        left_range = max(lefts) - min(lefts) if lefts else 0

        # Check for grid-like arrangement
        unique_tops = len(set(round(t / 100000) for t in tops))  # Group by approximate position
        unique_lefts = len(set(round(l / 100000) for l in lefts))

        if unique_tops >= 2 and unique_lefts >= 2:
            return {"score": 15, "type": "grid_layout"}
        elif top_range > 1000000 and left_range > 1000000:
            return {"score": 10, "type": "spread_layout"}
        else:
            return {"score": 0, "type": "linear_layout"}

    def _analyze_shape_variety(self, shapes):
        """Analyze variety in shape types and sizes"""
        if len(shapes) < 2:
            return 0

        shape_types = set()
        sizes = []

        for shape in shapes:
            shape_types.add(shape.get("type", "unknown"))
            pos = shape.get("position")
            if pos:
                size = pos["width"] * pos["height"]
                sizes.append(size)

        score = 0

        # Multiple shape types
        if len(shape_types) >= 3:
            score += 15
        elif len(shape_types) >= 2:
            score += 10

        # Consistent sizing (indicates process flow)
        if len(sizes) >= 3:
            avg_size = sum(sizes) / len(sizes)
            variations = [abs(size - avg_size) / avg_size for size in sizes if avg_size > 0]
            if variations and max(variations) < 0.5:  # Less than 50% variation
                score += 5

        return score

    def _analyze_text_density(self, text_blocks):
        """Analyze text characteristics for diagram indicators"""
        if not text_blocks:
            return 0

        short_text_count = 0
        total_blocks = len(text_blocks)

        for block in text_blocks:
            # Count average words per paragraph
            total_words = 0
            para_count = 0

            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    words = len(clean_text.split())
                    total_words += words
                    para_count += 1

            if para_count > 0:
                avg_words = total_words / para_count
                if avg_words <= 5:  # Short labels
                    short_text_count += 1

        # Score based on percentage of short text blocks
        if total_blocks > 0:
            short_ratio = short_text_count / total_blocks
            if short_ratio >= 0.7:  # 70%+ short text
                return 10
            elif short_ratio >= 0.5:  # 50%+ short text
                return 5

        return 0

    def _analyze_flow_patterns(self, shapes, lines, arrows, text_blocks):
        """Analyze for flow patterns and process keywords"""
        score = 0

        # Check for start/end keywords
        flow_keywords = ["start", "begin", "end", "finish", "process", "step", "decision"]
        action_words = ["create", "update", "check", "verify", "send", "receive", "analyze"]

        all_text = ""
        for block in text_blocks:
            for para in block.get("paragraphs", []):
                all_text += " " + para.get("clean_text", "").lower()

        flow_matches = sum(1 for keyword in flow_keywords if keyword in all_text)
        action_matches = sum(1 for keyword in action_words if keyword in all_text)

        if flow_matches >= 2:
            score += 20
        elif flow_matches >= 1:
            score += 10

        if action_matches >= 3:
            score += 10

        # Bonus for having both shapes and connecting elements
        if len(shapes) >= 3 and (len(lines) > 0 or len(arrows) > 0):
            score += 15

        return score

    def _analyze_negative_indicators(self, text_blocks, shapes):
        """Check for negative indicators that suggest NOT a diagram"""
        score = 0

        # Check for long paragraphs
        long_text_count = 0
        bullet_count = 0

        for block in text_blocks:
            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    word_count = len(clean_text.split())
                    if word_count > 20:  # Long paragraph
                        long_text_count += 1

                    # Check for bullet points
                    if para.get("hints", {}).get("is_bullet", False):
                        bullet_count += 1

        # Penalize long text
        if long_text_count >= 2:
            score -= 15

        # Penalize if mostly bullet points
        total_paras = sum(len(block.get("paragraphs", [])) for block in text_blocks)
        if total_paras > 0 and bullet_count / total_paras > 0.8:
            score -= 10

        # Penalize single column layout (all shapes vertically aligned)
        if len(shapes) >= 3:
            positions = [s.get("position") for s in shapes if s.get("position")]
            if len(positions) >= 3:
                lefts = [p["left"] for p in positions]
                left_variance = max(lefts) - min(lefts) if lefts else 0
                if left_variance < 500000:  # Very narrow horizontal spread
                    score -= 10

        return score

    def analyze_group_for_diagram(self, group_block):
        """
        Analyze a group to determine if it might be a diagram
        Returns analysis with confidence score
        """
        shapes = group_block.get("shapes", [])
        text_shapes = [s for s in shapes if s.get("type") == "text"]
        non_text_shapes = [s for s in shapes if s.get("type") != "text"]

        analysis = {
            "is_diagram": False,
            "type": "unknown",
            "confidence": 0,
            "shape_count": len(shapes),
            "reasons": []
        }

        # Quick exit for simple cases
        if len(shapes) < 3:
            return analysis

        # Gather all text content
        all_text = ""
        for shape in text_shapes:
            for para in shape.get("paragraphs", []):
                all_text += " " + para.get("clean_text", "")
        all_text = all_text.lower()

        confidence = 0

        # Check for diagram keywords
        flowchart_keywords = ["start", "end", "process", "decision", "flow", "workflow", "step"]
        org_keywords = ["manager", "director", "ceo", "team", "department", "reports to"]
        diagram_keywords = ["system", "component", "module", "architecture", "structure"]

        flowchart_matches = sum(1 for kw in flowchart_keywords if kw in all_text)
        org_matches = sum(1 for kw in org_keywords if kw in all_text)
        diagram_matches = sum(1 for kw in diagram_keywords if kw in all_text)

        if flowchart_matches >= 2:
            analysis["type"] = "flowchart"
            confidence += 40
            analysis["reasons"].append(f"Flowchart keywords: {flowchart_matches}")
        elif org_matches >= 2:
            analysis["type"] = "org_chart"
            confidence += 40
            analysis["reasons"].append(f"Org chart keywords: {org_matches}")
        elif diagram_matches >= 2:
            analysis["type"] = "diagram"
            confidence += 30
            analysis["reasons"].append(f"Diagram keywords: {diagram_matches}")

        # Check for non-text shapes (images, charts, etc.)
        if non_text_shapes:
            confidence += 20
            analysis["reasons"].append(f"Non-text shapes: {len(non_text_shapes)}")

        # Check for many shapes
        if len(shapes) >= 6:
            confidence += 15
            analysis["reasons"].append(f"Many shapes: {len(shapes)}")
        elif len(shapes) >= 4:
            confidence += 10

        # Check for complex positioning
        if self.has_complex_positioning_simple(shapes):
            confidence += 15
            analysis["reasons"].append("Complex positioning")

        analysis["confidence"] = confidence
        analysis["is_diagram"] = confidence >= 50  # Threshold for diagram detection

        return analysis

    def has_complex_positioning_simple(self, shapes):
        """Simple check for complex positioning"""
        positions = []
        for shape in shapes:
            pos = shape.get("position")
            if pos:
                positions.append((pos["top"], pos["left"]))

        if len(positions) < 3:
            return False

        # Check if shapes are spread out (not just stacked vertically)
        tops = [p[0] for p in positions]
        lefts = [p[1] for p in positions]

        top_range = max(tops) - min(tops) if tops else 0
        left_range = max(lefts) - min(lefts) if lefts else 0

        # If both horizontal and vertical spread, likely complex layout
        return top_range > 500000 and left_range > 1000000  # PowerPoint units

    def is_likely_slide_title(self, line):
        """
        Determine if a line is likely a slide title based on formatting and content.

        Args:
            line (str): The line to evaluate

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
            not text_content.endswith(('.', '!', '?', ';', ':')),  # Titles typically don't end with punctuation
            not self._contains_multiple_sentences(text_content),  # Titles are usually single phrases
            not text_content.lower().startswith(('the following', 'here are', 'this slide', 'key points')),
            # Avoid descriptive text
        ]

        # Additional positive indicators
        positive_indicators = [
            text_content.isupper(),  # All caps suggests title
            text_content.istitle(),  # Title case suggests title
            len(text_content.split()) <= 10,  # Short phrases are more likely titles
            any(word in text_content.lower() for word in
                ['overview', 'introduction', 'conclusion', 'agenda', 'objectives']),  # Common title words
        ]

        # Must meet basic criteria and have at least one positive indicator
        basic_criteria_met = all(title_indicators)
        has_positive_indicator = any(positive_indicators)

        return basic_criteria_met and (has_positive_indicator or len(text_content.split()) <= 6)

    def extract_title_from_bullet(self, bullet_line):
        """
        Extract clean title text from a bullet point line.

        Args:
            bullet_line (str): The bullet point line (e.g., "- Title Text")

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

    def convert_text_block_to_markdown(self, block):
        """Convert text block to markdown with proper formatting"""
        lines = []

        for para in block["paragraphs"]:
            line = self.convert_paragraph_to_markdown(para)
            if line:
                lines.append(line)

        # If entire shape is a hyperlink, wrap it
        result = "\n".join(lines)
        if block.get("shape_hyperlink") and result:
            result = f"[{result}]({block['shape_hyperlink']})"

        return result

    def convert_paragraph_to_markdown(self, para):
        """Convert paragraph to markdown with correct formatting"""
        if not para.get("clean_text"):
            return ""

        # Build formatted text from runs
        formatted_text = self.build_formatted_text_from_runs(para["formatted_runs"], para["clean_text"])

        # Now apply structural formatting based on hints
        hints = para.get("hints", {})

        # Bullets
        if hints.get("is_bullet", False):
            level = hints.get("bullet_level", 0)
            if level < 0:
                level = 0
            indent = "  " * level
            return f"{indent}- {formatted_text}"

        # Numbered lists
        elif hints.get("is_numbered", False):
            return f"1. {formatted_text}"

        # Headings
        elif hints.get("likely_heading", False):
            # Determine heading level
            if hints.get("all_caps") or len(para["clean_text"]) < 30:
                return f"## {formatted_text}"
            else:
                return f"### {formatted_text}"

        # Regular paragraph
        else:
            return formatted_text

    def build_formatted_text_from_runs(self, runs, clean_text):
        """Build formatted text from runs, handling edge cases"""
        if not runs:
            return clean_text

        # First check if we have any formatting at all
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

            # Apply formatting
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

    def convert_table_to_markdown(self, block):
        """Convert table to markdown"""
        if not block["data"]:
            return ""

        markdown = ""
        for i, row in enumerate(block["data"]):
            markdown += "| " + " | ".join(cell.replace("|", "\\|") for cell in row) + " |\n"

            # Add separator after header
            if i == 0:
                markdown += "| " + " | ".join("---" for _ in row) + " |\n"

        return markdown

    def convert_image_to_markdown(self, block):
        """Convert image to markdown"""
        image_md = f"![{block['alt_text']}](image)"

        if block.get("hyperlink"):
            image_md = f"[{image_md}]({block['hyperlink']})"

        return image_md

    def convert_chart_to_markdown(self, block):
        """Convert chart to markdown"""
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

        # Add comment for diagram conversion
        chart_md += f"\n<!-- DIAGRAM_CANDIDATE: chart, type={block.get('chart_type', 'unknown')} -->\n"

        if block.get("hyperlink"):
            chart_md = f"[{chart_md}]({block['hyperlink']})"

        return chart_md

    def convert_group_to_markdown(self, block):
        """Convert grouped shapes to markdown - handle all shape types"""
        # Get the extracted blocks from the group
        extracted_blocks = block.get("extracted_blocks", [])

        if not extracted_blocks:
            return ""

        # Convert each extracted block to markdown
        content_parts = []

        for extracted_block in extracted_blocks:
            if extracted_block["type"] == "text":
                content = self.convert_text_block_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "image":
                content = self.convert_image_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "table":
                content = self.convert_table_to_markdown(extracted_block)
                if content:
                    content_parts.append(content)
            elif extracted_block["type"] == "chart":
                content = self.convert_chart_to_markdown(extracted_block)
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

        # Join all content together
        group_md = "\n\n".join(content_parts) if content_parts else ""

        # Add shape-level hyperlink if present
        if block.get("hyperlink") and group_md:
            group_md = f"[{group_md}]({block['hyperlink']})"

        return group_md

    def is_actual_diagram(self, block, text_shapes_count, other_shapes_count):
        """
        Determine if a group represents an actual diagram or just grouped text
        Be very conservative - default to treating as grouped text
        """
        diagram_type = block.get("diagram_type", "text_group")

        # If it's identified as just a text group, definitely not a diagram
        if diagram_type == "text_group":
            return False

        # Only treat as diagram if we have very strong indicators
        if diagram_type in ["flowchart", "org_chart"]:
            # Even then, require additional evidence
            return other_shapes_count > 0 or text_shapes_count >= 5

        # For "diagram" type, require non-text shapes
        if diagram_type == "diagram":
            return other_shapes_count > 0

        # Default: treat as grouped text
        return False

    # Additional utility methods for the complete functionality

    def validate_file(self, file_path):
        """Validate that the file exists and is a supported PowerPoint format"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in self.supported_formats:
            raise ValueError(
                f"Unsupported file format: {file_ext}. Supported formats: {', '.join(self.supported_formats)}")

        return True

    def get_presentation_summary(self, presentation):
        """Get a quick summary of the presentation structure"""
        summary = {
            "total_slides": len(presentation.slides),
            "slide_details": []
        }

        for idx, slide in enumerate(presentation.slides, 1):
            slide_info = {
                "slide_number": idx,
                "shape_count": len(slide.shapes),
                "text_shapes": 0,
                "image_shapes": 0,
                "table_shapes": 0,
                "chart_shapes": 0,
                "group_shapes": 0
            }

            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    slide_info["image_shapes"] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    slide_info["table_shapes"] += 1
                elif hasattr(shape, 'has_chart') and shape.has_chart:
                    slide_info["chart_shapes"] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    slide_info["group_shapes"] += 1
                elif hasattr(shape, 'text_frame') or hasattr(shape, 'text'):
                    slide_info["text_shapes"] += 1

            summary["slide_details"].append(slide_info)

        return summary

    def extract_all_text(self, presentation):
        """Extract all text content from the presentation for text analysis"""
        all_text = []

        for slide_idx, slide in enumerate(presentation.slides, 1):
            slide_text = {
                "slide_number": slide_idx,
                "text_content": []
            }

            for shape in slide.shapes:
                text_content = self._extract_text_from_shape(shape)
                if text_content:
                    slide_text["text_content"].append(text_content)

            all_text.append(slide_text)

        return all_text

    def _extract_text_from_shape(self, shape):
        """Helper method to extract text from any shape type"""
        if hasattr(shape, 'text_frame') and shape.text_frame:
            text_parts = []
            for para in shape.text_frame.paragraphs:
                if para.text.strip():
                    text_parts.append(para.text.strip())
            return " ".join(text_parts) if text_parts else None

        elif hasattr(shape, 'text') and shape.text:
            return shape.text.strip() if shape.text.strip() else None

        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # Recursively extract text from grouped shapes
            group_text = []
            for child_shape in shape.shapes:
                child_text = self._extract_text_from_shape(child_shape)
                if child_text:
                    group_text.append(child_text)
            return " ".join(group_text) if group_text else None

        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            # Extract text from table cells
            table_text = []
            for row in shape.table.rows:
                row_text = []
                for cell in row.cells:
                    if hasattr(cell, 'text_frame') and cell.text_frame:
                        cell_content = []
                        for para in cell.text_frame.paragraphs:
                            if para.text.strip():
                                cell_content.append(para.text.strip())
                        if cell_content:
                            row_text.append(" ".join(cell_content))
                    elif cell.text and cell.text.strip():
                        row_text.append(cell.text.strip())
                if row_text:
                    table_text.append(" | ".join(row_text))
            return "\n".join(table_text) if table_text else None

        return None

    def export_to_json(self, presentation, file_path=None):
        """Export the extracted presentation data to JSON format"""
        structured_data = self.extract_presentation_data(presentation)
        metadata = self.extract_pptx_metadata(presentation, file_path or "unknown")

        export_data = {
            "metadata": metadata,
            "content": structured_data,
            "export_timestamp": datetime.now().isoformat(),
            "processor_version": "fixed_v3.0"
        }

        return json.dumps(export_data, indent=2, default=str)

    def process_file_complete(self, file_path, output_format="markdown", convert_slide_titles=True):
        """
        Complete file processing with multiple output options

        Args:
            file_path (str): Path to the PowerPoint file
            output_format (str): "markdown", "json", "text", or "summary"
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

        Returns:
            dict: Contains the processed content and metadata
        """
        # Validate file
        self.validate_file(file_path)

        # Load presentation
        prs = Presentation(file_path)

        # Extract all data
        metadata = self.extract_pptx_metadata(prs, file_path)
        structured_data = self.extract_presentation_data(prs)
        summary = self.get_presentation_summary(prs)
        all_text = self.extract_all_text(prs)

        result = {
            "file_path": file_path,
            "metadata": metadata,
            "summary": summary,
            "processing_timestamp": datetime.now().isoformat()
        }

        if output_format == "markdown":
            markdown = self.convert_structured_data_to_markdown(structured_data, convert_slide_titles)
            result["content"] = self.add_pptx_metadata_for_claude(markdown, metadata)
        elif output_format == "json":
            result["content"] = structured_data
            result["json_export"] = self.export_to_json(prs, file_path)
        elif output_format == "text":
            result["content"] = all_text
        elif output_format == "summary":
            result["content"] = {
                "summary": summary,
                "key_points": self._extract_key_points(all_text),
                "word_count": self._count_words(all_text)
            }
        else:
            raise ValueError(f"Unsupported output format: {output_format}")

        return result

    def _extract_key_points(self, all_text):
        """Extract potential key points from text content"""
        key_points = []

        for slide_text in all_text:
            for text_content in slide_text["text_content"]:
                if text_content:
                    # Look for bullet points or numbered lists
                    lines = text_content.split('\n')
                    for line in lines:
                        line = line.strip()
                        if (line and
                                (line.startswith(('•', '-', '*', '◦', '▪')) or
                                 re.match(r'^\d+[\.\)]\s', line) or
                                 len(line) < 100)):  # Short lines might be key points
                            key_points.append({
                                "slide": slide_text["slide_number"],
                                "text": line
                            })

        return key_points

    def _count_words(self, all_text):
        """Count total words in the presentation"""
        total_words = 0

        for slide_text in all_text:
            for text_content in slide_text["text_content"]:
                if text_content:
                    words = len(text_content.split())
                    total_words += words

        return total_words

    def debug_bullet_detection(self, file_path, slide_num=None, shape_num=None):
        """Debug bullet detection for specific shapes"""
        prs = Presentation(file_path)
        debug_info = []

        slides_to_check = [prs.slides[slide_num - 1]] if slide_num else prs.slides

        for slide_idx, slide in enumerate(slides_to_check, 1):
            shapes_to_check = [slide.shapes[shape_num]] if shape_num else slide.shapes

            for shape_idx, shape in enumerate(shapes_to_check):
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    shape_info = {
                        "slide": slide_idx if not slide_num else slide_num,
                        "shape": shape_idx,
                        "paragraphs": []
                    }

                    for para_idx, para in enumerate(shape.text_frame.paragraphs):
                        if para.text.strip():
                            para_info = {
                                "para_index": para_idx,
                                "text": para.text,
                                "powerpoint_level": getattr(para, 'level', None),
                                "xml_has_bullet": False,
                                "detected_as_bullet": False,
                                "final_level": -1
                            }

                            # Check XML
                            try:
                                if hasattr(para, '_p') and para._p is not None:
                                    xml_str = str(para._p.xml)
                                    if any(indicator in xml_str for indicator in ['buChar', 'buAutoNum', 'buFont']):
                                        para_info["xml_has_bullet"] = True
                            except:
                                pass

                            # Process with fixed method
                            processed = self.process_paragraph_fixed(para)
                            if processed:
                                para_info["detected_as_bullet"] = processed["hints"]["is_bullet"]
                                para_info["final_level"] = processed["hints"]["bullet_level"]

                            shape_info["paragraphs"].append(para_info)

                    if shape_info["paragraphs"]:
                        debug_info.append(shape_info)

        return debug_info


# Convenience functions for backward compatibility and ease of use

def convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles=True):
    """
    Convenience function to maintain backward compatibility

    Args:
        file_path (str): Path to the PowerPoint file
        convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings
    """
    processor = PowerPointProcessor()
    return processor.convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles)


def process_powerpoint_file(file_path, output_format="markdown", convert_slide_titles=True):
    """
    Convenience function for complete file processing

    Args:
        file_path (str): Path to the PowerPoint file
        output_format (str): "markdown", "json", "text", or "summary"
        convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

    Returns:
        dict: Processed content and metadata
    """
    processor = PowerPointProcessor()
    return processor.process_file_complete(file_path, output_format, convert_slide_titles)


def debug_bullets(file_path, slide_num=None, shape_num=None):
    """
    Debug bullet detection

    Args:
        file_path (str): Path to the PowerPoint file
        slide_num (int, optional): Specific slide number to debug
        shape_num (int, optional): Specific shape number to debug

    Returns:
        list: Debug information for each shape
    """
    processor = PowerPointProcessor()
    return processor.debug_bullet_detection(file_path, slide_num, shape_num)