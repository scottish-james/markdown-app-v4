"""
PowerPoint Processor - Main Coordinator Class
XML-first approach with MarkItDown fallback when XML unavailable.
Simplified: Use XML when available, MarkItDown when not.

ARCHITECTURE OVERVIEW:
This is the main orchestrator that coordinates all PowerPoint processing components.
It implements a dual-strategy approach: sophisticated XML-based processing when
possible, with graceful fallback to MarkItDown for simpler conversion.

PROCESSING STRATEGIES:
1. XML-first Strategy: Full component pipeline with rich feature extraction
   - AccessibilityOrderExtractor for reading order
   - ContentExtractor with TextProcessor for detailed formatting
   - DiagramAnalyzer for diagram detection
   - MetadataExtractor for document context
   - MarkdownConverter for final output

2. MarkItDown Fallback: Simple, reliable conversion when XML unavailable
   - Direct file conversion using MarkItDown library
   - Basic metadata annotation
   - No advanced features but guaranteed compatibility

DECISION LOGIC:
The processor automatically determines which strategy to use based on XML
accessibility. This provides optimal results when possible while ensuring
compatibility across all PowerPoint versions and file conditions.

COMPONENT COORDINATION:
Acts as the central coordinator that:
- Initializes all component instances
- Manages processing flow between components
- Handles errors and fallbacks gracefully
- Provides unified API for external consumers
- Maintains processing state and debugging information
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
    Main PowerPoint processor implementing dual-strategy processing architecture.

    COMPONENT ORCHESTRATION:
    Coordinates six specialized components in a pipeline:
    1. AccessibilityOrderExtractor: Determines proper reading order
    2. ContentExtractor: Extracts content from different shape types
    3. TextProcessor: Handles advanced text formatting
    4. DiagramAnalyzer: Identifies potential diagrams
    5. MarkdownConverter: Generates final markdown output
    6. MetadataExtractor: Provides document context

    PROCESSING MODES:
    - Sophisticated XML Mode: Full feature pipeline when XML accessible
    - MarkItDown Fallback Mode: Simple conversion when XML unavailable
    - Automatic Detection: Chooses mode based on XML accessibility

    ERROR HANDLING STRATEGY:
    - Component isolation: Individual component failures don't crash pipeline
    - Graceful degradation: Falls back to simpler processing on errors
    - Comprehensive logging: Tracks processing methods and issues
    - Recovery mechanisms: Multiple fallback options at each stage
    """

    def __init__(self, use_accessibility_order=True):
        """
        Initialize the PowerPoint processor with all component dependencies.

        COMPONENT INITIALIZATION:
        Creates instances of all specialized processors with appropriate
        configuration. Components are designed to be stateless for
        thread safety and reusability.

        CONFIGURATION OPTIONS:
        - use_accessibility_order: Controls reading order strategy
        - Component-specific settings passed through to individual processors

        DEPENDENCY MANAGEMENT:
        - All components are initialized upfront for performance
        - MarkItDown initialized for fallback processing
        - Component configuration is centralized here

        Args:
            use_accessibility_order (bool): Whether to use XML-based accessibility reading order
        """
        self.use_accessibility_order = use_accessibility_order

        # Initialize specialized components for XML-based processing
        # Each component handles a specific aspect of PowerPoint processing
        self.accessibility_extractor = AccessibilityOrderExtractor(use_accessibility_order)
        self.content_extractor = ContentExtractor()
        self.diagram_analyzer = DiagramAnalyzer()
        self.text_processor = TextProcessor()
        self.markdown_converter = MarkdownConverter()
        self.metadata_extractor = MetadataExtractor()

        # Initialize MarkItDown for fallback processing
        # Provides simple, reliable conversion when XML processing fails
        self.markitdown = MarkItDown()

        # Supported file format configuration
        self.supported_formats = ['.pptx', '.ppt']

    def convert_pptx_to_markdown_enhanced(self, file_path, convert_slide_titles=True):
        """
        Main entry point implementing XML-first processing with MarkItDown fallback.

        PROCESSING ALGORITHM:
        1. Check XML accessibility for the PowerPoint file
        2. If XML available: Use sophisticated component pipeline
        3. If XML unavailable: Fall back to MarkItDown conversion
        4. Handle errors gracefully with meaningful error messages

        XML ACCESSIBILITY CHECK:
        Determines processing strategy by attempting to access PowerPoint's
        internal XML structure. This check is fast and non-destructive.

        STRATEGY SELECTION:
        - XML Available: Full feature extraction with all components
        - XML Unavailable: Simple text extraction with basic formatting
        - Error Handling: Comprehensive error messages for debugging

        FEATURE AVAILABILITY:
        XML Mode: All features (accessibility order, diagram analysis, metadata)
        Fallback Mode: Basic text extraction only

        Args:
            file_path (str): Path to the PowerPoint file to process
            convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

        Returns:
            str: Enhanced markdown content with all available features

        Raises:
            Exception: With descriptive error message if processing completely fails
        """
        try:
            # Determine processing strategy based on XML accessibility
            if self._has_xml_access(file_path):
                return self._sophisticated_xml_processing(file_path, convert_slide_titles)
            else:
                return self._simple_markitdown_processing(file_path)

        except Exception as e:
            raise Exception(f"Error processing PowerPoint file: {str(e)}")

    def _has_xml_access(self, file_path):
        """
        Check if XML-based processing is possible for the given file.

        XML ACCESSIBILITY FACTORS:
        1. File format compatibility (python-pptx support)
        2. File corruption or damage
        3. Password protection or encryption
        4. Version-specific XML structure differences

        TESTING APPROACH:
        1. Load presentation with python-pptx
        2. Access first slide if available
        3. Test XML accessibility through AccessibilityOrderExtractor
        4. Return boolean result for strategy selection

        PERFORMANCE CONSIDERATIONS:
        - Fast check using minimal file access
        - Fails quickly for incompatible files
        - No heavy processing during detection
        - Caches presentation loading when possible

        ERROR HANDLING:
        All exceptions indicate XML inaccessibility, leading to
        MarkItDown fallback processing.

        Args:
            file_path (str): Path to the PowerPoint file

        Returns:
            bool: True if XML processing possible, False otherwise
        """
        try:
            # Attempt to load presentation with python-pptx
            prs = Presentation(file_path)

            # Test XML accessibility using first slide
            if len(prs.slides) > 0:
                first_slide = prs.slides[0]
                return self.accessibility_extractor._has_xml_access(first_slide)

            # No slides to test - assume XML unavailable
            return False
        except Exception:
            # Any loading failure means XML processing impossible
            return False

    def _sophisticated_xml_processing(self, file_path, convert_slide_titles):
        """
        Execute full-featured processing pipeline when XML is accessible.

        SOPHISTICATED PROCESSING PIPELINE:
        1. Load PowerPoint presentation with python-pptx
        2. Extract comprehensive metadata for document context
        3. Process all slides through component pipeline
        4. Convert structured data to markdown
        5. Add metadata context for Claude AI processing
        6. Perform diagram analysis and add results
        7. Return enhanced markdown with all features

        COMPONENT INTEGRATION:
        - MetadataExtractor: Document context and properties
        - AccessibilityOrderExtractor: Proper reading order
        - ContentExtractor + TextProcessor: Rich content extraction
        - MarkdownConverter: Clean markdown generation
        - DiagramAnalyzer: Diagram detection and scoring

        FEATURE RICHNESS:
        This mode provides all available features:
        - Accessibility-aware reading order
        - Advanced text formatting preservation
        - Diagram detection and analysis
        - Comprehensive metadata inclusion
        - Slide title post-processing

        PERFORMANCE CHARACTERISTICS:
        More CPU and memory intensive than fallback mode but provides
        significantly richer feature set and better output quality.

        Args:
            file_path (str): Path to the PowerPoint file
            convert_slide_titles (bool): Whether to convert slide titles

        Returns:
            str: Enhanced markdown with all sophisticated features
        """
        print("🎯 Using sophisticated XML-based processing...")

        # Load presentation for full processing
        prs = Presentation(file_path)

        # Extract comprehensive metadata for document context
        pptx_metadata = self.metadata_extractor.extract_pptx_metadata(prs, file_path)

        # Process entire presentation through component pipeline
        structured_data = self.extract_presentation_data(prs)

        # Convert structured data to clean markdown
        markdown = self.markdown_converter.convert_structured_data_to_markdown(
            structured_data, convert_slide_titles
        )

        # Enhance with metadata context for Claude AI processing
        markdown_with_metadata = self.metadata_extractor.add_pptx_metadata_for_claude(
            markdown, pptx_metadata
        )

        # Add diagram analysis results if diagrams detected
        diagram_analysis = self.diagram_analyzer.analyze_structured_data_for_diagrams(structured_data)
        if diagram_analysis:
            markdown_with_metadata += "\n\n" + diagram_analysis

        return markdown_with_metadata

    def _simple_markitdown_processing(self, file_path):
        """
        Execute simple fallback processing using MarkItDown library.

        MARKITDOWN FALLBACK STRATEGY:
        When XML processing is impossible (corrupted files, unsupported
        formats, etc.), fall back to MarkItDown for basic text extraction.

        FALLBACK CAPABILITIES:
        - Basic text extraction from slides
        - Simple formatting preservation
        - Image placeholder detection
        - Table structure recognition (limited)
        - No advanced features (accessibility order, diagram analysis)

        ERROR HANDLING:
        MarkItDown has different failure modes than python-pptx.
        Handles multiple attribute access patterns for result object.

        RESULT OBJECT VARIATIONS:
        Different MarkItDown versions expose results differently:
        - result.markdown (preferred)
        - result.text_content (alternative)
        - Provides helpful error messages for debugging

        METADATA ANNOTATION:
        Adds simple comment indicating fallback mode was used
        for debugging and processing transparency.

        Args:
            file_path (str): Path to the PowerPoint file

        Returns:
            str: Basic markdown content from MarkItDown

        Raises:
            Exception: If MarkItDown processing also fails
        """
        print("📄 XML not available - using MarkItDown fallback...")

        try:
            # Use MarkItDown library for simple conversion
            result = self.markitdown.convert(file_path)

            # Handle different result object formats across MarkItDown versions
            try:
                markdown_content = result.markdown
            except AttributeError:
                try:
                    markdown_content = result.text_content
                except AttributeError:
                    raise Exception("Neither 'markdown' nor 'text_content' attribute found on result object")

            # Add processing method annotation for debugging
            metadata_comment = f"\n<!-- Converted using MarkItDown fallback - XML not available -->\n"

            return metadata_comment + markdown_content

        except Exception as e:
            raise Exception(f"MarkItDown processing failed: {str(e)}")

    def extract_presentation_data(self, presentation):
        """
        Extract structured data from entire presentation using component coordination.

        PRESENTATION-LEVEL PROCESSING:
        Coordinates extraction across all slides while maintaining
        presentation structure and metadata.

        DATA STRUCTURE:
        Creates comprehensive structured representation:
        - total_slides: Overall presentation metrics
        - slides: Array of slide data with extraction metadata
        - Each slide contains content blocks and processing information

        SLIDE PROCESSING:
        Delegates individual slide processing to extract_slide_data()
        while maintaining presentation-level context and ordering.

        METADATA PRESERVATION:
        Tracks processing method and slide count for downstream
        analysis and debugging purposes.

        SCALABILITY:
        Processes slides sequentially to manage memory usage
        while maintaining slide order and relationships.

        Args:
            presentation: python-pptx Presentation object

        Returns:
            dict: Structured presentation data with all slides processed
        """
        data = {
            "total_slides": len(presentation.slides),
            "slides": []
        }

        # Process each slide individually while maintaining order
        for slide_idx, slide in enumerate(presentation.slides, 1):
            slide_data = self.extract_slide_data(slide, slide_idx)
            data["slides"].append(slide_data)

        return data

    def extract_slide_data(self, slide, slide_number):
        """
        Extract content from individual slide using coordinated component pipeline.

        SLIDE PROCESSING PIPELINE:
        1. AccessibilityOrderExtractor: Determine proper shape reading order
        2. ContentExtractor: Extract content from each shape in order
        3. TextProcessor: Handle text formatting (called by ContentExtractor)
        4. Return structured slide data with extraction metadata

        READING ORDER IMPORTANCE:
        Proper reading order is crucial for:
        - Accessibility compliance
        - Logical content flow
        - Accurate content understanding
        - Proper slide title detection

        SHAPE PROCESSING:
        Each shape processed individually through ContentExtractor
        which routes to appropriate handlers based on shape type.

        EXTRACTION METADATA:
        Tracks extraction method used for debugging and optimization:
        - semantic_accessibility_order: Full XML-based ordering
        - positional_order: Fallback positioning
        - markitdown_fallback: Simple enumeration

        Args:
            slide: python-pptx Slide object
            slide_number (int): Slide number (1-based) for debugging

        Returns:
            dict: Slide data with content blocks and extraction metadata
        """
        # Get shapes in proper reading order using AccessibilityOrderExtractor
        ordered_shapes = self.accessibility_extractor.get_slide_reading_order(slide, slide_number)

        slide_data = {
            "slide_number": slide_number,
            "content_blocks": [],
            "extraction_method": self.accessibility_extractor.get_last_extraction_method()
        }

        # Extract content from each shape using ContentExtractor + TextProcessor
        for shape in ordered_shapes:
            block = self.content_extractor.extract_shape_content(shape, self.text_processor)
            if block:
                slide_data["content_blocks"].append(block)

        return slide_data

    def debug_accessibility_order(self, file_path, slide_number=1):
        """
        Debug method for analyzing reading order extraction and processing decisions.

        DEBUGGING PURPOSE:
        Provides detailed insight into:
        - XML accessibility determination
        - Reading order extraction methods
        - Shape processing results
        - Component decision making

        DEBUGGING OUTPUT:
        - Processing strategy selection reasoning
        - Shape count and order information
        - Extraction method identification
        - Sample shape content preview

        SLIDE SELECTION:
        Focuses on single slide for detailed analysis without
        overwhelming output. Slide 1 is default as it often
        contains representative content.

        ERROR HANDLING:
        Comprehensive error handling with helpful error messages
        for debugging processing failures.

        DEVELOPMENT TOOL:
        Intended for development and troubleshooting, not
        production use. Provides verbose diagnostic output.

        Args:
            file_path (str): Path to PowerPoint file
            slide_number (int): Slide number to debug (1-based indexing)
        """
        try:
            # First check and report XML accessibility
            if not self._has_xml_access(file_path):
                print(f"❌ XML not available for {file_path}")
                print("Would use MarkItDown fallback in production")
                return

            # XML available - proceed with sophisticated processing debug
            prs = Presentation(file_path)
            if slide_number > len(prs.slides):
                print(f"Slide {slide_number} not found. Presentation has {len(prs.slides)} slides.")
                return

            slide = prs.slides[slide_number - 1]
            print(f"🎯 XML available - debugging sophisticated processing...")

            # Debug accessibility order extraction in detail
            print(f"\n=== DEBUGGING SLIDE {slide_number} READING ORDER ===")
            print(f"Total shapes: {len(slide.shapes)}")

            # Test and report accessibility extraction results
            ordered_shapes = self.accessibility_extractor.get_slide_reading_order(slide, slide_number)
            print(f"✅ Extraction method: {self.accessibility_extractor.get_last_extraction_method()}")
            print(f"✅ Ordered shapes: {len(ordered_shapes)}")

            # Show sample shape information for verification
            print("\n🎯 SHAPE ORDER:")
            for i, shape in enumerate(ordered_shapes[:5]):  # Show first 5 shapes only
                shape_type = str(shape.shape_type).split('.')[-1]
                text_preview = ""
                try:
                    # Extract text preview for shape identification
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
        Get comprehensive processing summary without performing full conversion.

        SUMMARY PURPOSE:
        Provides quick assessment of processing capabilities and
        expected results without the overhead of full conversion.

        INFORMATION INCLUDED:
        - XML accessibility determination
        - Processing strategy selection
        - Feature availability assessment
        - Basic presentation metrics
        - Preview of processing results

        XML MODE ANALYSIS:
        When XML is available, provides detailed preview:
        - Slide count and structure
        - Extraction method for each slide
        - Shape count and content type detection
        - Processing method verification

        FALLBACK MODE ANALYSIS:
        When XML unavailable, provides basic information:
        - Fallback processing notification
        - Limited feature availability
        - Simple processing expectations

        USAGE SCENARIOS:
        - Pre-processing validation
        - Processing capability assessment
        - Debugging and troubleshooting
        - User interface information display

        Args:
            file_path (str): Path to PowerPoint file to analyze

        Returns:
            dict: Comprehensive processing summary with capabilities and preview
        """
        try:
            # Determine and report XML accessibility
            has_xml = self._has_xml_access(file_path)

            summary = {
                "file_path": file_path,
                "has_xml_access": has_xml,
                "processing_method": "sophisticated_xml" if has_xml else "markitdown_fallback"
            }

            if has_xml:
                # Detailed analysis for XML-capable processing
                prs = Presentation(file_path)

                summary.update({
                    "slide_count": len(prs.slides),
                    "extraction_method": "accessibility_order" if self.use_accessibility_order else "positional",
                    "has_diagram_analysis": True,
                    "slides_preview": []
                })

                # Generate preview for first few slides
                for i, slide in enumerate(prs.slides[:3], 1):  # Preview first 3 slides
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
                # Basic summary for MarkItDown fallback processing
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
        Configure reading order extraction method for all processing.

        CONFIGURATION SCOPE:
        Updates configuration for current processor instance:
        - Main processor setting
        - AccessibilityOrderExtractor configuration
        - Affects all subsequent processing operations

        EXTRACTION METHODS:
        - True: Use XML-based accessibility order (sophisticated)
        - False: Use positional order (simple, faster)

        RUNTIME RECONFIGURATION:
        Allows changing extraction method without creating new
        processor instance, useful for testing and optimization.

        PROPAGATION:
        Configuration change propagates to AccessibilityOrderExtractor
        to ensure consistent behavior across the processing pipeline.

        Args:
            use_accessibility_order (bool): Whether to use XML-based accessibility order
        """
        self.use_accessibility_order = use_accessibility_order
        self.accessibility_extractor.use_accessibility_order = use_accessibility_order


# Convenience functions for backward compatibility and simple usage
def convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles=True):
    """
    Convenience function maintaining backward compatibility with simple API.

    BACKWARD COMPATIBILITY:
    Maintains existing API for users who don't need advanced configuration
    or component-level control.

    SIMPLE USAGE PATTERN:
    Provides one-line conversion for most common use case:
    convert_pptx_to_markdown_enhanced("presentation.pptx")

    CONFIGURATION:
    Uses default configuration appropriate for most use cases:
    - Accessibility order enabled
    - Slide title conversion enabled
    - All components enabled

    Args:
        file_path (str): Path to the PowerPoint file
        convert_slide_titles (bool): Whether to convert slide titles from bullets to H1 headings

    Returns:
        str: Enhanced markdown content with all available features
    """
    processor = PowerPointProcessor()
    return processor.convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles)


def process_powerpoint_file(file_path, output_format="markdown", convert_slide_titles=True):
    """
    Convenience function for comprehensive file processing with multiple output options.

    OUTPUT FORMAT OPTIONS:
    - "markdown": Standard markdown output (default)
    - "json": Structured data output
    - "text": Plain text output
    - "summary": Processing summary without full conversion

    ENHANCED FUNCTIONALITY:
    Provides additional metadata and processing information
    beyond simple markdown conversion.

    METADATA INCLUSION:
    When XML processing available, includes comprehensive
    PowerPoint metadata in the result structure.

    PROCESSING TRANSPARENCY:
    Result includes processing method information for
    debugging and quality assessment.

    Args:
        file_path (str): Path to the PowerPoint file
        output_format (str): Desired output format
        convert_slide_titles (bool): Whether to convert slide titles

    Returns:
        dict: Processed content with metadata and processing information
    """
    processor = PowerPointProcessor()

    if output_format == "summary":
        # Return processing summary without full conversion
        return processor.get_processing_summary(file_path)
    else:
        # Perform full processing and return enhanced result
        markdown_content = processor.convert_pptx_to_markdown_enhanced(file_path, convert_slide_titles)

        result = {
            "content": markdown_content,
            "format": output_format,
            "processing_method": "sophisticated_xml" if processor._has_xml_access(file_path) else "markitdown_fallback"
        }

        # Add comprehensive metadata if XML processing was used
        if processor._has_xml_access(file_path):
            try:
                prs = Presentation(file_path)
                result["metadata"] = processor.metadata_extractor.extract_pptx_metadata(prs, file_path)
            except Exception:
                # Metadata extraction failed - continue without metadata
                pass

        return result

