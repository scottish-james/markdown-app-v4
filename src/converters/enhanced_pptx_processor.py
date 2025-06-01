"""
PowerPoint Processor - Fixed and Complete
Maintains all original functionality while fixing bullet detection
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
    """Complete PowerPoint processing with fixed bullet detection"""

    def __init__(self):
        self.supported_formats = ['.pptx', '.ppt']

    def convert_pptx_to_markdown_enhanced(self, file_path):
        """
        Main entry point: Convert PowerPoint to structured data, then to markdown with embedded metadata
        """
        try:
            prs = Presentation(file_path)

            # Extract PowerPoint metadata first
            pptx_metadata = self.extract_pptx_metadata(prs, file_path)

            # Extract structured data
            structured_data = self.extract_presentation_data(prs)

            # Convert to basic markdown
            markdown = self.convert_structured_data_to_markdown(structured_data)

            # Add PowerPoint metadata as comments for Claude to use
            markdown_with_metadata = self.add_pptx_metadata_for_claude(markdown, pptx_metadata)

            return markdown_with_metadata
        except Exception as e:
            raise Exception(f"Error processing PowerPoint file: {str(e)}")

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
        """Extract shape content with proper type detection"""
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return self.extract_image(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            return self.extract_table(shape.table)
        elif hasattr(shape, 'has_chart') and shape.has_chart:
            return self.extract_chart(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return self.extract_group(shape)
        elif hasattr(shape, 'text_frame') and shape.text_frame:
            return self.extract_text_frame_fixed(shape.text_frame, shape)
        elif hasattr(shape, 'text') and shape.text:
            return self.extract_plain_text(shape)
        return None

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

    def extract_runs(self, runs):
        """Extract formatting and hyperlinks from runs"""
        formatted_runs = []

        for run in runs:
            run_data = {
                "text": run.text,
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

            formatted_runs.append(run_data)

        return formatted_runs

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
        """Extract content from grouped shapes"""
        try:
            group_data = {
                "type": "group",
                "shapes": [],
                "connections": [],
                "hyperlink": self.extract_shape_hyperlink(shape)
            }

            # Process shapes in the group
            for child_shape in shape.shapes:
                child_data = self.extract_shape_content(child_shape)
                if child_data:
                    # Add position information for diagram analysis
                    if hasattr(child_shape, 'top') and hasattr(child_shape, 'left'):
                        child_data["position"] = {
                            "top": child_shape.top,
                            "left": child_shape.left,
                            "width": getattr(child_shape, 'width', 0),
                            "height": getattr(child_shape, 'height', 0)
                        }
                    group_data["shapes"].append(child_data)

            # Analyze for potential diagram patterns
            group_data["diagram_type"] = self.analyze_diagram_pattern(group_data["shapes"])

            return group_data

        except Exception:
            return None

    def analyze_diagram_pattern(self, shapes):
        """Analyze shapes to determine if they form a recognizable diagram pattern"""
        if not shapes:
            return "unknown"

        text_shapes = [s for s in shapes if s.get("type") == "text"]

        # Look for keywords to identify diagram type
        flowchart_keywords = ["start", "end", "process", "decision", "flow"]
        org_keywords = ["manager", "director", "ceo", "team", "department"]

        all_text = " ".join([
            " ".join([p.get("clean_text", "") for p in shape.get("paragraphs", [])])
            for shape in text_shapes
        ]).lower()

        if any(keyword in all_text for keyword in flowchart_keywords):
            return "flowchart"
        elif any(keyword in all_text for keyword in org_keywords):
            return "org_chart"
        elif len(shapes) >= 3:
            return "diagram"

        return "unknown"

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

    def convert_structured_data_to_markdown(self, data):
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

        return "\n\n".join(filter(None, markdown_parts))

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
        """Convert grouped shapes to markdown with diagram analysis"""
        diagram_type = block.get("diagram_type", "unknown")

        # Start with diagram identification
        group_md = f"**Diagram ({diagram_type})**\n\n"

        # Convert individual shapes
        shape_content = []
        for shape in block.get("shapes", []):
            if shape.get("type") == "text":
                content = self.convert_text_block_to_markdown(shape)
                if content:
                    shape_content.append(content)
            elif shape.get("type") == "image":
                image_md = self.convert_image_to_markdown(shape)
                if image_md:
                    shape_content.append(image_md)

        if shape_content:
            group_md += "\n".join(shape_content)

        # Add diagram conversion hint
        group_md += f"\n\n<!-- DIAGRAM_CANDIDATE: {diagram_type}, shapes={len(block.get('shapes', []))} -->\n"

        if block.get("hyperlink"):
            group_md = f"[{group_md}]({block['hyperlink']})"

        return group_md

    # Additional utility methods for the complete superfile functionality

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
            "processor_version": "fixed_v2.0"
        }

        return json.dumps(export_data, indent=2, default=str)

    def process_file_complete(self, file_path, output_format="markdown"):
        """
        Complete file processing with multiple output options

        Args:
            file_path (str): Path to the PowerPoint file
            output_format (str): "markdown", "json", "text", or "summary"

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
            markdown = self.convert_structured_data_to_markdown(structured_data)
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

def convert_pptx_to_markdown_enhanced(file_path):
    """
    Convenience function to maintain backward compatibility
    """
    processor = PowerPointProcessor()
    return processor.convert_pptx_to_markdown_enhanced(file_path)


def process_powerpoint_file(file_path, output_format="markdown"):
    """
    Convenience function for complete file processing

    Args:
        file_path (str): Path to the PowerPoint file
        output_format (str): "markdown", "json", "text", or "summary"

    Returns:
        dict: Processed content and metadata
    """
    processor = PowerPointProcessor()
    return processor.process_file_complete(file_path, output_format)


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