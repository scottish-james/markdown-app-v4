"""
Metadata Extractor - Handles PowerPoint metadata extraction and formatting
Extracts comprehensive document metadata and formats it for Claude enhancement
"""

import os
from datetime import datetime


class MetadataExtractor:
    """
    Extracts and formats PowerPoint metadata for document enhancement.
    Handles core properties, document properties, and file information.
    """

    def extract_pptx_metadata(self, presentation, file_path):
        """
        Extract comprehensive metadata from PowerPoint file.

        Args:
            presentation: python-pptx Presentation object
            file_path (str): Path to the PowerPoint file

        Returns:
            dict: Comprehensive metadata dictionary
        """
        metadata = {}

        try:
            # Core properties from PowerPoint
            core_props = presentation.core_properties

            # Basic file information
            metadata['filename'] = os.path.basename(file_path)
            metadata['file_size'] = os.path.getsize(file_path) if os.path.exists(file_path) else None

            # Document properties
            metadata.update(self._extract_document_properties(core_props))

            # Date information
            metadata.update(self._extract_date_properties(core_props))

            # Revision and tracking
            metadata.update(self._extract_revision_properties(core_props))

            # Presentation-specific information
            metadata.update(self._extract_presentation_properties(presentation))

            # Application information
            metadata.update(self._extract_application_properties(presentation))

        except Exception as e:
            print(f"Warning: Could not extract some metadata: {e}")

        return metadata

    def _extract_document_properties(self, core_props):
        """
        Extract basic document properties.

        Args:
            core_props: PowerPoint core properties object

        Returns:
            dict: Document properties
        """
        return {
            'title': getattr(core_props, 'title', '') or '',
            'author': getattr(core_props, 'author', '') or '',
            'subject': getattr(core_props, 'subject', '') or '',
            'keywords': getattr(core_props, 'keywords', '') or '',
            'comments': getattr(core_props, 'comments', '') or '',
            'category': getattr(core_props, 'category', '') or '',
            'content_status': getattr(core_props, 'content_status', '') or '',
            'language': getattr(core_props, 'language', '') or '',
            'version': getattr(core_props, 'version', '') or '',
        }

    def _extract_date_properties(self, core_props):
        """
        Extract date-related properties.

        Args:
            core_props: PowerPoint core properties object

        Returns:
            dict: Date properties
        """
        return {
            'created': getattr(core_props, 'created', None),
            'modified': getattr(core_props, 'modified', None),
            'last_modified_by': getattr(core_props, 'last_modified_by', '') or '',
            'last_printed': getattr(core_props, 'last_printed', None),
        }

    def _extract_revision_properties(self, core_props):
        """
        Extract revision and identifier properties.

        Args:
            core_props: PowerPoint core properties object

        Returns:
            dict: Revision properties
        """
        return {
            'revision': getattr(core_props, 'revision', None),
            'identifier': getattr(core_props, 'identifier', '') or '',
        }

    def _extract_presentation_properties(self, presentation):
        """
        Extract presentation-specific properties.

        Args:
            presentation: python-pptx Presentation object

        Returns:
            dict: Presentation properties
        """
        metadata = {
            'slide_count': len(presentation.slides)
        }

        # Extract slide master and layout information
        try:
            slide_masters = presentation.slide_masters
            if slide_masters:
                metadata['slide_master_count'] = len(slide_masters)

                # Get layout names
                layout_names = []
                for master in slide_masters:
                    for layout in master.slide_layouts:
                        if hasattr(layout, 'name') and layout.name:
                            layout_names.append(layout.name)

                metadata['layout_types'] = ', '.join(set(layout_names)) if layout_names else ''
            else:
                metadata['slide_master_count'] = 0
                metadata['layout_types'] = ''
        except Exception:
            metadata['slide_master_count'] = 0
            metadata['layout_types'] = ''

        return metadata

    def _extract_application_properties(self, presentation):
        """
        Extract application-related properties.

        Args:
            presentation: python-pptx Presentation object

        Returns:
            dict: Application properties
        """
        metadata = {
            'application': '',
            'app_version': '',
            'company': '',
            'doc_security': None
        }

        try:
            app_props = presentation.app_properties if hasattr(presentation, 'app_properties') else None
            if app_props:
                metadata['application'] = getattr(app_props, 'application', '') or ''
                metadata['app_version'] = getattr(app_props, 'app_version', '') or ''
                metadata['company'] = getattr(app_props, 'company', '') or ''
                metadata['doc_security'] = getattr(app_props, 'doc_security', None)
        except Exception:
            pass

        return metadata

    def add_pptx_metadata_for_claude(self, markdown_content, metadata):
        """
        Add PowerPoint metadata as comments for Claude to incorporate.

        Args:
            markdown_content (str): Original markdown content
            metadata (dict): Extracted metadata

        Returns:
            str: Markdown content with embedded metadata
        """
        # Format metadata for Claude
        metadata_comments = "\n<!-- POWERPOINT METADATA FOR CLAUDE:\n"

        # Add document information
        metadata_comments += self._format_document_metadata(metadata)

        # Add date information
        metadata_comments += self._format_date_metadata(metadata)

        # Add file information
        metadata_comments += self._format_file_metadata(metadata)

        # Add presentation information
        metadata_comments += self._format_presentation_metadata(metadata)

        metadata_comments += "-->\n"

        # Add metadata at the beginning
        return metadata_comments + markdown_content

    def _format_document_metadata(self, metadata):
        """
        Format document-related metadata for Claude.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            str: Formatted document metadata
        """
        formatted = ""

        if metadata.get('title'):
            formatted += f"Document Title: {metadata['title']}\n"
        if metadata.get('author'):
            formatted += f"Author: {metadata['author']}\n"
        if metadata.get('subject'):
            formatted += f"Subject: {metadata['subject']}\n"
        if metadata.get('keywords'):
            formatted += f"Keywords: {metadata['keywords']}\n"
        if metadata.get('category'):
            formatted += f"Category: {metadata['category']}\n"
        if metadata.get('comments'):
            formatted += f"Document Comments: {metadata['comments']}\n"
        if metadata.get('content_status'):
            formatted += f"Content Status: {metadata['content_status']}\n"
        if metadata.get('language'):
            formatted += f"Language: {metadata['language']}\n"
        if metadata.get('version'):
            formatted += f"Version: {metadata['version']}\n"

        return formatted

    def _format_date_metadata(self, metadata):
        """
        Format date-related metadata for Claude.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            str: Formatted date metadata
        """
        formatted = ""

        if metadata.get('created'):
            formatted += f"Created Date: {metadata['created']}\n"
        if metadata.get('modified'):
            formatted += f"Last Modified: {metadata['modified']}\n"
        if metadata.get('last_modified_by'):
            formatted += f"Last Modified By: {metadata['last_modified_by']}\n"
        if metadata.get('last_printed'):
            formatted += f"Last Printed: {metadata['last_printed']}\n"

        return formatted

    def _format_file_metadata(self, metadata):
        """
        Format file-related metadata for Claude.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            str: Formatted file metadata
        """
        formatted = ""

        formatted += f"Filename: {metadata.get('filename', 'unknown')}\n"

        if metadata.get('file_size'):
            file_size_mb = metadata['file_size'] / (1024 * 1024)
            formatted += f"File Size: {file_size_mb:.2f} MB\n"

        if metadata.get('application'):
            formatted += f"Created With: {metadata['application']}\n"
        if metadata.get('company'):
            formatted += f"Company: {metadata['company']}\n"

        return formatted

    def _format_presentation_metadata(self, metadata):
        """
        Format presentation-specific metadata for Claude.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            str: Formatted presentation metadata
        """
        formatted = ""

        formatted += f"Slide Count: {metadata.get('slide_count', 0)}\n"

        if metadata.get('slide_master_count'):
            formatted += f"Slide Masters: {metadata['slide_master_count']}\n"
        if metadata.get('layout_types'):
            formatted += f"Layout Types: {metadata['layout_types']}\n"

        return formatted

    def get_metadata_summary(self, metadata):
        """
        Get a human-readable summary of the metadata.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            dict: Summary information
        """
        summary = {
            'has_title': bool(metadata.get('title')),
            'has_author': bool(metadata.get('author')),
            'slide_count': metadata.get('slide_count', 0),
            'file_size_mb': None,
            'creation_date': metadata.get('created'),
            'last_modified': metadata.get('modified'),
            'has_keywords': bool(metadata.get('keywords')),
            'application': metadata.get('application', 'Unknown'),
        }

        if metadata.get('file_size'):
            summary['file_size_mb'] = round(metadata['file_size'] / (1024 * 1024), 2)

        return summary

    def validate_metadata(self, metadata):
        """
        Validate metadata completeness and quality.

        Args:
            metadata (dict): Metadata dictionary

        Returns:
            dict: Validation results
        """
        validation = {
            'completeness_score': 0,
            'issues': [],
            'recommendations': []
        }

        # Check for essential fields
        essential_fields = ['title', 'author', 'slide_count']
        present_fields = sum(1 for field in essential_fields if metadata.get(field))
        validation['completeness_score'] = (present_fields / len(essential_fields)) * 100

        # Check for issues
        if not metadata.get('title'):
            validation['issues'].append("No document title")
            validation['recommendations'].append("Add a descriptive title to the presentation")

        if not metadata.get('author'):
            validation['issues'].append("No author information")
            validation['recommendations'].append("Set author information in document properties")

        if metadata.get('slide_count', 0) == 0:
            validation['issues'].append("No slides detected")

        if not metadata.get('keywords'):
            validation['recommendations'].append("Add keywords to improve searchability")

        return validation