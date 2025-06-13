"""
PowerPoint Processing Components
Modular architecture for PowerPoint to Markdown conversion
"""

# Import all component classes
from .powerpoint_processor import PowerPointProcessor
from .accessibility_extractor import AccessibilityOrderExtractor
from .content_extractor import ContentExtractor
from .text_processor import TextProcessor
from .diagram_analyzer import DiagramAnalyzer
from .markdown_converter import MarkdownConverter
from .metadata_extractor import MetadataExtractor

# Import convenience functions
from .powerpoint_processor import (
    convert_pptx_to_markdown_enhanced,
    process_powerpoint_file
)

# Export everything for easy importing
__all__ = [
    # Main processor class
    'PowerPointProcessor',

    # Component classes
    'AccessibilityOrderExtractor',
    'ContentExtractor',
    'TextProcessor',
    'DiagramAnalyzer',
    'MarkdownConverter',
    'MetadataExtractor',

    # Convenience functions
    'convert_pptx_to_markdown_enhanced',
    'process_powerpoint_file'
]

# Version info
__version__ = "2.0.0"
__architecture__ = "modular"