"""
Simplified Enterprise LLM Converter for PowerPoint Processing
Uses proven Claude prompts with enterprise SageMaker endpoints
"""

import os
import json
import requests
import logging
from typing import Tuple, Optional, Dict
import time

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Proven prompts from working Claude system
PPTX_PROCESSING_PROMPT = """
You are a PowerPoint to Markdown converter. You receive roughly converted markdown from PowerPoint slides and must clean it up into professional, well-structured markdown for AI training and vector database storage.

Your job:
1. Fix bullet point hierarchies - create proper nested lists with 2-space indentation
2. Identify slide titles and format as # headers
3. Identify sub heading and format as ### headers 
4. Preserve ALL hyperlinks and formatting (bold, italic)
5. Fix broken list structures
6. Ensure tables are properly formatted
7. Clean up spacing and structure
8. USE POWERPOINT METADATA when available (look for HTML comments with "POWERPOINT METADATA FOR CLAUDE")
9. Split all slides with ---
10. ADD COMPREHENSIVE METADATA at the end for vector database optimization

Key Rules:
- Keep ALL original text content - this is critical as this is for regulatory documents
- Fix bullet nesting based on context and content
- Make short, standalone text into appropriate ### headers 
- Preserve all hyperlinks exactly as provided
- Use proper markdown syntax throughout
- Look for numbered sequences that are out of order and fix them
- INCORPORATE PowerPoint metadata (author, creation date, etc.) into the final metadata section

CRITICAL: Always end with a metadata section that includes:
- TLDR/Executive Summary (2-3 sentences)
- Key Topics/Themes (for embeddings)
- Target Audience (inferred)
- Key Concepts/Terms
- File metadata (author, creation date, version, etc. from PowerPoint properties)
- Diagram Types (list any Mermaid diagrams created)

The input will have slide markers like `<!-- Slide 1 -->` and may include PowerPoint metadata in HTML comments and diagram candidates.

Format the metadata section like this at the very end:

---
## DOCUMENT METADATA (for AI/Vector DB)

**TLDR:** [2-3 sentence summary of the entire presentation]

**Key Topics:** [comma-separated list of main topics/themes]

**Content Type:** [e.g., Educational, Business Presentation, Training Material, etc.]

**Related Topics:** [concepts that would be complementary to search for]

**Visual Elements:** [number and types of diagrams, charts, images converted to Mermaid]

**File Properties:**
- **Author:** [from PowerPoint metadata]
- **Created Date:** [from PowerPoint metadata]
- **Last Modified:** [from PowerPoint metadata]
- **Version:** [from PowerPoint metadata]
- **Company/Organization:** [from PowerPoint metadata]
- **Document Title:** [from PowerPoint metadata]
- **Keywords:** [from PowerPoint metadata]
- **Category:** [from PowerPoint metadata]
- **Slide Dimensions:** [from PowerPoint metadata]
---

Output clean, readable markdown that maintains the original document's intent but fixes obvious ordering issues, converts diagrams to Mermaid, and adds rich metadata for AI training.
"""

DOCUMENT_PROCESSING_PROMPT = """
You are a markdown formatting expert. Your task is to:
1. Fix any syntax errors in the markdown
2. Improve the structure and hierarchy of headers
3. Ensure consistent formatting throughout
4. Enhance bullet points and numbered lists
5. Properly format tables and code blocks
6. Add appropriate spacing between sections
7. Maintain the original content without adding new information
8. Preserve all links and references

Return ONLY the enhanced markdown content without any explanations or additional text.
"""


class EnterpriseLLMClient:
    """
    Simplified client for enterprise LLM hosted on SageMaker
    """

    def __init__(self):
        """Initialize the enterprise LLM client"""
        self.jwt_token = self._load_jwt_token()
        self.model_url = self._load_model_url()
        self.headers = {
            "Authorization": f"Bearer {self.jwt_token}",
            "Content-Type": "application/json"
        }

    def _load_jwt_token(self) -> str:
        """Load JWT token from file"""
        try:
            with open("JWT_token.txt", "r") as f:
                token = f.read().strip()
            if not token:
                raise ValueError("JWT token file is empty")
            return token
        except FileNotFoundError:
            raise ValueError("JWT_token.txt file not found")
        except Exception as e:
            raise ValueError(f"Error reading JWT token: {str(e)}")

    def _load_model_url(self) -> str:
        """Load model URL from file - simplified to use single endpoint"""
        try:
            with open("model_url.txt", "r") as f:
                content = f.read().strip()

            if not content:
                raise ValueError("Model URL file is empty")

            # Support both single URL and JSON format
            if content.startswith('{'):
                urls = json.loads(content)
                # Use content model URL or first available
                return urls.get("content", list(urls.values())[0])
            else:
                return content

        except FileNotFoundError:
            raise ValueError("model_url.txt file not found")
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON format in model_url.txt")
        except Exception as e:
            raise ValueError(f"Error reading model URL: {str(e)}")

    def call_model(self, prompt: str, content: str, max_tokens: int = 4096) -> Tuple[str, Optional[str]]:
        """
        Call the enterprise LLM with retry logic

        Args:
            prompt (str): System prompt
            content (str): User content
            max_tokens (int): Maximum tokens for response

        Returns:
            Tuple[str, Optional[str]]: Response content and error message
        """
        # Prepare request payload
        payload = {
            "messages": [
                {"role": "system", "content": prompt},
                {"role": "user", "content": content}
            ],
            "max_tokens": max_tokens,
            "temperature": 0.1
        }

        # Retry logic for enterprise stability
        max_retries = 3
        for attempt in range(max_retries):
            try:
                logger.info(f"Calling enterprise model (attempt {attempt + 1}/{max_retries})")

                response = requests.post(
                    self.model_url,
                    headers=self.headers,
                    json=payload,
                    timeout=120  # 2 minute timeout
                )

                if response.status_code == 200:
                    result = response.json()

                    # Extract response based on common SageMaker response formats
                    if "choices" in result and result["choices"]:
                        enhanced_content = result["choices"][0]["message"]["content"]
                    elif "generated_text" in result:
                        enhanced_content = result["generated_text"]
                    elif "content" in result:
                        enhanced_content = result["content"]
                    else:
                        enhanced_content = str(result)

                    logger.info("Successfully processed content with enterprise model")
                    return enhanced_content, None

                elif response.status_code == 429:  # Rate limited
                    wait_time = 2 ** attempt  # Exponential backoff
                    logger.warning(f"Rate limited, waiting {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue

                else:
                    error_msg = f"API error {response.status_code}: {response.text}"
                    logger.error(error_msg)
                    if attempt == max_retries - 1:
                        return content, error_msg

            except requests.exceptions.Timeout:
                error_msg = f"Request timeout on attempt {attempt + 1}"
                logger.error(error_msg)
                if attempt == max_retries - 1:
                    return content, error_msg

            except Exception as e:
                error_msg = f"Request failed: {str(e)}"
                logger.error(error_msg)
                if attempt == max_retries - 1:
                    return content, error_msg

        return content, "All retry attempts failed"


class EnterpriseLLMEnhancer:
    """
    Simplified PowerPoint processor using enterprise LLM
    """

    def __init__(self):
        """Initialize the enterprise LLM enhancer"""
        self.client = EnterpriseLLMClient()

    def enhance_powerpoint_content(self, structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
    Tuple[str, Optional[str]]:
        """
        Simple, direct processing method using proven prompts

        Args:
            structured_data (dict): Structured presentation data
            metadata (dict): PowerPoint metadata
            source_filename (str): Source filename

        Returns:
            Tuple[str, Optional[str]]: Enhanced content and error message
        """
        logger.info(f"Starting enterprise LLM processing for {source_filename}")

        try:
            # Convert structured data to basic markdown
            basic_markdown = self._convert_structured_to_basic_markdown(structured_data, metadata, source_filename)

            # Determine content type for prompt selection
            if "pptx" in source_filename.lower() or "ppt" in source_filename.lower():
                prompt = PPTX_PROCESSING_PROMPT
                user_content = f"""
                Please clean up this PowerPoint markdown conversion for AI training and vector database storage:

                **Source:** {source_filename}

                **Content to clean up:**
                {basic_markdown}

IMPORTANT: 
1. Fix the structure, bullet hierarchies, and formatting while preserving all content and hyperlinks
2. If you see content with ordering words like "first", "second", "third", "fourth", "fifth" etc., REORDER those items into the correct sequence
3. ADD COMPREHENSIVE METADATA at the end following the specified format for vector database optimization

The metadata section is CRITICAL for AI training - make sure to analyze the content thoroughly and provide rich, searchable metadata that will help with embeddings and retrieval."""

            else:
                prompt = DOCUMENT_PROCESSING_PROMPT
                user_content = f"""Please enhance this markdown document:

**Source Information:**
- Filename: {source_filename}

**Original Markdown Content:**
{basic_markdown}

Please apply the formatting standards and ensure all content is preserved while improving the structure, consistency, and readability."""

            # Process with enterprise LLM
            enhanced_content, error = self.client.call_model(prompt, user_content)

            if error:
                logger.error(f"Enterprise LLM processing error: {error}")
                return basic_markdown, error

            logger.info(f"Enterprise LLM processing completed for {source_filename}")
            return enhanced_content, None

        except Exception as e:
            error_msg = f"Enterprise LLM enhancement failed: {str(e)}"
            logger.error(error_msg)
            return self._convert_structured_to_basic_markdown(structured_data, metadata, source_filename), error_msg

    def _convert_structured_to_basic_markdown(self, structured_data: Dict, metadata: Dict, source_filename: str) -> str:
        """
        Convert structured data to basic markdown for processing
        """
        markdown_parts = []

        # Add metadata as HTML comment for the LLM
        if metadata:
            markdown_parts.append("<!-- POWERPOINT METADATA FOR CLAUDE:")
            for key, value in metadata.items():
                if value:
                    markdown_parts.append(f"{key}: {value}")
            markdown_parts.append("-->")

        # Process slides
        for slide in structured_data.get("slides", []):
            markdown_parts.append(f"\n<!-- Slide {slide['slide_number']} -->\n")

            # Process content blocks
            for block in slide.get("content_blocks", []):
                if block.get("type") == "text":
                    for para in block.get("paragraphs", []):
                        text = para.get("clean_text", "").strip()
                        if text:
                            hints = para.get("hints", {})
                            if hints.get("is_bullet"):
                                level = hints.get("bullet_level", 0)
                                indent = "  " * level
                                markdown_parts.append(f"{indent}- {text}")
                            else:
                                markdown_parts.append(text)

                elif block.get("type") == "table":
                    # Basic table conversion
                    table_data = block.get("data", [])
                    if table_data:
                        for i, row in enumerate(table_data):
                            markdown_parts.append("| " + " | ".join(row) + " |")
                            if i == 0:  # Header separator
                                markdown_parts.append("| " + " | ".join("---" for _ in row) + " |")

                elif block.get("type") == "image":
                    alt_text = block.get("alt_text", "Image")
                    markdown_parts.append(f"![{alt_text}](image)")

                elif block.get("type") == "chart":
                    title = block.get("title", "Chart")
                    markdown_parts.append(f"**Chart: {title}**")
                    markdown_parts.append("<!-- DIAGRAM_CANDIDATE: chart -->")

        return "\n\n".join(filter(None, markdown_parts))


def enhance_markdown_with_enterprise_llm(structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
Tuple[str, Optional[str]]:
    """
    Convenience function to enhance PowerPoint content using enterprise LLM

    Args:
        structured_data (dict): Structured presentation data
        metadata (dict): PowerPoint metadata
        source_filename (str): Source filename

    Returns:
        Tuple[str, Optional[str]]: Enhanced content and error message
    """
    try:
        enhancer = EnterpriseLLMEnhancer()
        return enhancer.enhance_powerpoint_content(structured_data, metadata, source_filename)
    except Exception as e:
        error_msg = f"Enterprise LLM enhancement failed: {str(e)}"
        logger.error(error_msg)

        # Return basic markdown as fallback
        fallback_content = f"# {source_filename}\n\n"
        fallback_content += "**Error:** Enterprise LLM processing failed, showing original content.\n\n"

        for slide in structured_data.get("slides", []):
            fallback_content += f"## Slide {slide['slide_number']}\n\n"
            for block in slide.get("content_blocks", []):
                if block.get("type") == "text":
                    for para in block.get("paragraphs", []):
                        text = para.get("clean_text", "").strip()
                        if text:
                            fallback_content += f"{text}\n\n"

        return fallback_content, error_msg


# Integration helper for existing codebase
class EnterpriseLLMMarkdownEnhancer:
    """
    Drop-in replacement for Claude enhancer that maintains the same interface
    """

    def __init__(self):
        self.enhancer = EnterpriseLLMEnhancer()

    def enhance_markdown(self, structured_data: Dict, metadata: Dict, source_filename: str = "unknown",
                         content_type: str = "PowerPoint") -> Tuple[str, Optional[str]]:
        """
        Main interface method that matches the Claude enhancer signature

        Args:
            structured_data (dict): Structured data instead of raw markdown
            metadata (dict): Document metadata
            source_filename (str): Source filename
            content_type (str): Content type (maintained for compatibility)

        Returns:
            Tuple[str, Optional[str]]: Enhanced content and error message
        """
        return self.enhancer.enhance_powerpoint_content(structured_data, metadata, source_filename)