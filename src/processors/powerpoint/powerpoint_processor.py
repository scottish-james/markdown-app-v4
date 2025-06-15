"""
PowerPoint Processor - Main Coordinator Class
XML-first approach with MarkItDown fallback when XML unavailable.
Simplified: Use XML when available, MarkItDown when not.
"""

from pptx import Presentation
import os
from datetime import datetime
from markitdown import MarkItDown

from .accessibility_extractor import AccessibilityOrderExtractor
from .content_extractor import ContentExtractor
from .diagram_analyzer import DiagramAnalyzer
from .text_processor import TextProcessor
from .markdown_converter import MarkdownConverter
from .metadata_extractor import MetadataExtractor


class PowerPointProcessor:
    """
    Main PowerPoint processor that uses XML when available, MarkItDown when not.
    This class determines which processing path to use based on XML availability.
    """

    def __init__(self, use_accessibility_order=True):
        """
        Initialize the PowerPoint processor with all its components.

        Args:
            use_accessibility_order (bool): Whether to use accessibility reading order
        """
        self.use_accessibility_order = use_accessibility_order

        # Initialize specialized components (only used when XML available)
        self.accessibility_extractor = AccessibilityOrderExtractor(use_accessibility_order)
        self.content_extractor = ContentExtractor()
        self.diagram_analyzer = DiagramAnalyzer()
        self.text_processor = TextProcessor()
        self.markdown_converter = MarkdownConverter()
        self.metadata_extractor = MetadataExtractor()

        # Initialize MarkItDown for fallback
        self.markitdown = MarkItDown()

        self.supported_formats = ['.pptx', '.ppt']

    def convert_pptx_to_markdown_enhanced(self, file_path, convert_slide_titles=True):
        """
        Main entry point: XML-first approach with MarkItDown fallback.

        Args:
            file_path (str): Path to the PowerPoint file
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

        Returns:
            str: Enhanced markdown content
        """
        try:
            # Check if we have XML access
            if self._has_xml_access(file_path):
                return self._sophisticated_xml_processing(file_path, convert_slide_titles)
            else:
                return self._simple_markitdown_processing(file_path)

        except Exception as e:
            raise Exception(f"Error processing PowerPoint file: {str(e)}")

    def _has_xml_access(self, file_path):
        """
        Check if we can access XML data from the PowerPoint file.

        Args:
            file_path (str): Path to the PowerPoint file

        Returns:
            bool: True if XML is accessible
        """
        try:
            # Try to load presentation and access XML
            prs = Presentation(file_path)

            # Check if we can access slide XML
            if len(prs.slides) > 0:
                first_slide = prs.slides[0]
                return self.accessibility_extractor._has_xml_access(first_slide)

            return False
        except Exception:
            return False

    def _sophisticated_xml_processing(self, file_path, convert_slide_titles):
        """
        Use sophisticated processing when XML is available.

        Args:
            file_path (str): Path to the PowerPoint file
            convert_slide_titles (bool): Whether to convert slide titles

        Returns:
            str: Enhanced markdown with metadata and diagram analysis
        """
        print("🎯 Using sophisticated XML-based processing...")

        # Load presentation
        prs = Presentation(file_path)

        # Extract metadata
        pptx_metadata = self.metadata_extractor.extract_pptx_metadata(prs, file_path)

        # Extract structured data using all components
        structured_data = self.extract_presentation_data(prs)

        # Convert to basic markdown
        markdown = self.markdown_converter.convert_structured_data_to_markdown(
            structured_data, convert_slide_titles
        )

        # Add metadata for Claude enhancement
        markdown_with_metadata = self.metadata_extractor.add_pptx_metadata_for_claude(
            markdown, pptx_metadata
        )

        # Add diagram analysis
        diagram_analysis = self.diagram_analyzer.analyze_structured_data_for_diagrams(structured_data)
        if diagram_analysis:
            markdown_with_metadata += "\n\n" + diagram_analysis

        return markdown_with_metadata

    def _simple_markitdown_processing(self, file_path):
        """
        Use simple MarkItDown processing when XML is not available.

        Args:
            file_path (str): Path to the PowerPoint file

        Returns:
            str: Basic markdown content from MarkItDown
        """
        print("📄 XML not available - using MarkItDown fallback...")

        try:
            # Use MarkItDown for conversion
            result = self.markitdown.convert(file_path)

            # Extract content
            try:
                markdown_content = result.markdown
            except AttributeError:
                try:
                    markdown_content = result.text_content
                except AttributeError:
                    raise Exception("Neither 'markdown' nor 'text_content' attribute found on result object")

            # Add simple metadata comment
            metadata_comment = f"\n<!-- Converted using MarkItDown fallback - XML not available -->\n"

            return metadata_comment + markdown_content

        except Exception as e:
            raise Exception(f"MarkItDown processing failed: {str(e)}")

    def extract_presentation_data(self, presentation):
        """
        Extract all content from the presentation using coordinated components.
        Only called when XML is available.

        Args:
            presentation: python-pptx Presentation object

        Returns:
            dict: Structured presentation data
        """
        data = {
            "total_slides": len(presentation.slides),
            "slides": []
        }

        for slide_idx, slide in enumerate(presentation.slides, 1):
            slide_data = self.extract_slide_data(slide, slide_idx)
            data["slides"].append(slide_data)

        return data

    def extract_slide_data(self, slide, slide_number):
        """
        Extract content from a single slide using accessibility order.
        Only called when XML is available.

        Args:
            slide: python-pptx Slide object
            slide_number (int): Slide number (1-based)

        Returns:
            dict: Slide data with content blocks
        """
        # Get shapes in the correct reading order
        ordered_shapes = self.accessibility_extractor.get_slide_reading_order(slide, slide_number)

        slide_data = {
            "slide_number": slide_number,
            "content_blocks": [],
            "extraction_method": self.accessibility_extractor.get_last_extraction_method()
        }

        # Extract content from each shape
        for shape in ordered_shapes:
            block = self.content_extractor.extract_shape_content(shape, self.text_processor)
            if block:
                slide_data["content_blocks"].append(block)

        return slide_data

    def debug_accessibility_order(self, file_path, slide_number=1):
        """
        Debug method to analyze reading order extraction.

        Args:
            file_path (str): Path to PowerPoint file
            slide_number (int): Slide number to debug (1-based)
        """
        try:
            # Check XML availability first
            if not self._has_xml_access(file_path):
                print(f"❌ XML not available for {file_path}")
                print("Would use MarkItDown fallback in production")
                return

            prs = Presentation(file_path)
            if slide_number > len(prs.slides):
                print(f"Slide {slide_number} not found. Presentation has {len(prs.slides)} slides.")
                return

            slide = prs.slides[slide_number - 1]
            print(f"🎯 XML available - debugging sophisticated processing...")

            # Debug the accessibility extractor
            print(f"\n=== DEBUGGING SLIDE {slide_number} READING ORDER ===")
            print(f"Total shapes: {len(slide.shapes)}")

            # Test accessibility order extraction
            ordered_shapes = self.accessibility_extractor.get_slide_reading_order(slide, slide_number)
            print(f"✅ Extraction method: {self.accessibility_extractor.get_last_extraction_method()}")
            print(f"✅ Ordered shapes: {len(ordered_shapes)}")

            print("\n🎯 SHAPE ORDER:")
            for i, shape in enumerate(ordered_shapes[:5]):  # Show first 5
                shape_type = str(shape.shape_type).split('.')[-1]
                text_preview = ""
                try:
                    if hasattr(shape, 'text') and shape.text:
                        text_preview = shape.text.strip()[:40] + "..."
                    elif hasattr(shape, 'text_frame') and shape.text_frame:
                        text_preview = shape.text_frame.text.strip()[:40] + "..."
                except:
                    text_preview = "No text"

                print(f"  {i + 1}. [{shape_type}] {text_preview}")

        except Exception as e:
            print(f"Debug failed: {str(e)}")

    def get_processing_summary(self, file_path):
        """
        Get a summary of what will be processed without full conversion.

        Args:
            file_path (str): Path to PowerPoint file

        Returns:
            dict: Processing summary information
        """
        try:
            # Check XML availability
            has_xml = self._has_xml_access(file_path)

            summary = {
                "file_path": file_path,
                "has_xml_access": has_xml,
                "processing_method": "sophisticated_xml" if has_xml else "markitdown_fallback"
            }

            if has_xml:
                # Detailed analysis for XML cases
                prs = Presentation(file_path)

                summary.update({
                    "slide_count": len(prs.slides),
                    "extraction_method": "accessibility_order" if self.use_accessibility_order else "positional",
                    "has_diagram_analysis": True,
                    "slides_preview": []
                })

                # Preview first few slides
                for i, slide in enumerate(prs.slides[:3], 1):
                    ordered_shapes = self.accessibility_extractor.get_slide_reading_order(slide, i)

                    slide_preview = {
                        "slide_number": i,
                        "shape_count": len(ordered_shapes),
                        "has_text": any(hasattr(shape, 'text_frame') and shape.text_frame
                                        for shape in ordered_shapes),
                        "extraction_method": self.accessibility_extractor.get_last_extraction_method()
                    }
                    summary["slides_preview"].append(slide_preview)
            else:
                # Basic summary for MarkItDown cases
                summary.update({
                    "slide_count": "unknown",
                    "extraction_method": "markitdown_fallback",
                    "has_diagram_analysis": False,
                    "note": "XML not available - using simple MarkItDown conversion"
                })

            return summary

        except Exception as e:
            return {"error": str(e)}

    def configure_extraction_method(self, use_accessibility_order):
        """
        Configure the extraction method for reading order.

        Args:
            use_accessibility_order (bool): Whether to use accessibility order
        """
        self.use_accessibility_order = use_accessibility_order
        self.accessibility_extractor.use_accessibility_order = use_accessibility_order


# Convenience functions for backward compatibility
def convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles=True):
    """
    Convenience function to maintain backward compatibility.

    Args:
        file_path (str): Path to the PowerPoint file
        convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

    Returns:
        str: Enhanced markdown content
    """
    processor = PowerPointProcessor()
    return processor.convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles)


def process_powerpoint_file(file_path, output_format="markdown", convert_slide_titles=True):
    """
    Convenience function for complete file processing.

    Args:
        file_path (str): Path to the PowerPoint file
        output_format (str): "markdown", "json", "text", or "summary"
        convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

    Returns:
        dict: Processed content and metadata
    """
    processor = PowerPointProcessor()

    if output_format == "summary":
        return processor.get_processing_summary(file_path)
    else:
        # Process the file
        markdown_content = processor.convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles)

        result = {
            "content": markdown_content,
            "format": output_format,
            "processing_method": "sophisticated_xml" if processor._has_xml_access(file_path) else "markitdown_fallback"
        }

        # Add metadata if XML was available
        if processor._has_xml_access(file_path):
            try:
                prs = Presentation(file_path)
                result["metadata"] = processor.metadata_extractor.extract_pptx_metadata(prs, file_path)
            except Exception:
                pass

        return result

