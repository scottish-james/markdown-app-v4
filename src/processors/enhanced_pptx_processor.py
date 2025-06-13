"""
PowerPoint Processor - Backward Compatibility Facade
Maintained for existing imports while using new modular architecture.

This file replaces the original 1,500+ line monolithic enhanced_pptx_processor.py
with a clean facade that delegates to the new modular components.
"""

# Import from PowerPoint module
from .powerpoint import (
    PowerPointProcessor,
    convert_pptx_to_markdown_enhanced,
    process_powerpoint_file,
    AccessibilityOrderExtractor,
    ContentExtractor,
    TextProcessor,
    DiagramAnalyzer,
    MarkdownConverter,
    MetadataExtractor
)

# Re-export everything for backward compatibility
__all__ = [
    # Main processor class
    'PowerPointProcessor',

    # Convenience functions (maintain exact same API)
    'convert_pptx_to_markdown_enhanced',
    'process_powerpoint_file',

    # Component classes (for advanced usage)
    'AccessibilityOrderExtractor',
    'ContentExtractor',
    'TextProcessor',
    'DiagramAnalyzer',
    'MarkdownConverter',
    'MetadataExtractor'
]


# Convenience function for debugging (new feature)
def debug_powerpoint_reading_order(file_path, slide_number=1):
    """
    Debug reading order extraction for a specific slide.

    Args:
        file_path (str): Path to PowerPoint file
        slide_number (int): Slide number to debug (1-based)
    """
    processor = PowerPointProcessor()
    processor.debug_accessibility_order(file_path, slide_number)


# Convenience function for processing summary (new feature)
def get_powerpoint_processing_summary(file_path):
    """
    Get a summary of what will be processed without full conversion.

    Args:
        file_path (str): Path to PowerPoint file

    Returns:
        dict: Processing summary information
    """
    processor = PowerPointProcessor()
    return processor.get_processing_summary(file_path)


# Add new functions to exports
__all__.extend([
    'debug_powerpoint_reading_order',
    'get_powerpoint_processing_summary'
])

# Version information
__version__ = "2.0.0"
__architecture__ = "modular"

# Module docstring for users
__doc__ = """
PowerPoint to Markdown Processor (Modular Architecture v2.0)

This module provides PowerPoint to Markdown conversion with:
- Accessibility-based reading order extraction
- Advanced text formatting and bullet detection  
- v19 diagram analysis and scoring
- Comprehensive metadata extraction
- Clean, testable modular architecture

Basic Usage:
    from src.processors.enhanced_pptx_processor import convert_pptx_to_markdown_enhanced

    markdown = convert_pptx_to_markdown_enhanced("presentation.pptx")

Advanced Usage:
    from src.processors.enhanced_pptx_processor import PowerPointProcessor

    processor = PowerPointProcessor(use_accessibility_order=True)
    summary = processor.get_processing_summary("presentation.pptx")
    markdown = processor.convert_pptx_to_markdown_enhanced("presentation.pptx")

New in v2.0:
- Modular component architecture
- Enhanced debugging capabilities
- Processing summaries and validation
- Better error handling and isolation
- Component-level testing support
"""