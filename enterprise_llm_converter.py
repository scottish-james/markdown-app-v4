"""
Enterprise LLM Converter for PowerPoint Processing
Replaces Claude integration with enterprise SageMaker LLM endpoints
"""

import os
import json
import requests
import logging
from typing import Tuple, Optional, List, Dict
import time

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Enhanced system prompts for different content types
METADATA_PROCESSING_PROMPT = """
You are a document metadata analyser. You receive PowerPoint metadata and must create a comprehensive summary for document management and search optimisation.

Your job:
1. Analyse the metadata comprehensively
2. Create executive summary and key insights
3. Identify document purpose and audience
4. Suggest categorisation and tags
5. Highlight any data quality issues
6. Format everything in clear, professional UK English

Output clean, actionable metadata analysis that helps with document discovery and management.
"""

CONTENT_PROCESSING_PROMPT = """
You are a PowerPoint content processor. You receive slide content and must clean it up into professional, well-structured markdown for business use.

Your job:
1. Fix bullet point hierarchies - create proper nested lists with 2-space indentation
2. Ensure titles are properly formatted as headers
3. Preserve ALL content and hyperlinks
4. Fix broken list structures and improve readability
5. Clean up spacing and structure
6. Maintain professional UK English throughout

Key Rules:
- Keep ALL original content - this is critical for business documents
- Use proper markdown syntax throughout
- Preserve all hyperlinks exactly as provided
- Ensure consistent formatting and structure

Output clean, readable markdown that maintains the original document's intent whilst improving structure and readability.
"""

DIAGRAM_PROCESSING_PROMPT = """
You are a diagram and visual content specialist. You receive slide content that contains potential diagrams and charts.

Your job:
1. Identify visual elements that could be enhanced
2. Create Mermaid diagram representations where appropriate
3. Improve descriptions of charts and visual elements
4. Suggest how visual content could be better represented in markdown
5. Maintain all original content whilst enhancing visual descriptions

Output enhanced markdown with improved visual content representation.
"""


class EnterpriseLLMClient:
    """
    Client for enterprise LLM hosted on SageMaker with multiple model routing
    """

    def __init__(self):
        """
        Initialise the enterprise LLM client
        """
        self.jwt_token = self._load_jwt_token()
        self.model_urls = self._load_model_urls()
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

    def _load_model_urls(self) -> Dict[str, str]:
        """Load model URLs from file - supporting multiple endpoints"""
        try:
            with open("model_url.txt", "r") as f:
                content = f.read().strip()

            # Support both single URL and JSON format for multiple models
            if content.startswith('{'):
                # JSON format: {"metadata": "url1", "content": "url2", "diagram": "url3"}
                urls = json.loads(content)
            else:
                # Single URL format - use for all model types
                urls = {
                    "metadata": content,
                    "content": content,
                    "diagram": content
                }

            # Validate required URLs
            required_types = ["metadata", "content", "diagram"]
            for model_type in required_types:
                if model_type not in urls:
                    logger.warning(f"No URL specified for {model_type} model, using default")
                    urls[model_type] = list(urls.values())[0]  # Use first available URL

            return urls

        except FileNotFoundError:
            raise ValueError("model_url.txt file not found")
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON format in model_url.txt")
        except Exception as e:
            raise ValueError(f"Error reading model URLs: {str(e)}")

    def call_model(self, prompt: str, content: str, model_type: str = "content", max_tokens: int = 4096) -> Tuple[
        str, Optional[str]]:
        """
        Call the enterprise LLM with retry logic

        Args:
            prompt (str): System prompt
            content (str): User content
            model_type (str): Type of model to use ('metadata', 'content', 'diagram')
            max_tokens (int): Maximum tokens for response

        Returns:
            Tuple[str, Optional[str]]: Response content and error message
        """
        if model_type not in self.model_urls:
            return content, f"Unknown model type: {model_type}"

        url = self.model_urls[model_type]

        # Prepare request payload (adjust based on your SageMaker endpoint format)
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
                logger.info(f"Calling {model_type} model (attempt {attempt + 1}/{max_retries})")

                response = requests.post(
                    url,
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
                        # Fallback - return the whole response as string
                        enhanced_content = str(result)

                    logger.info(f"Successfully processed content with {model_type} model")
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
    Enhanced PowerPoint processor using enterprise LLM with intelligent routing
    """

    def __init__(self):
        """
        Initialise the enterprise LLM enhancer
        """
        self.client = EnterpriseLLMClient()

    def enhance_powerpoint_content(self, structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
    Tuple[str, Optional[str]]:
        """
        Main processing method that routes different content types to appropriate models

        Args:
            structured_data (dict): Structured presentation data
            metadata (dict): PowerPoint metadata
            source_filename (str): Source filename

        Returns:
            Tuple[str, Optional[str]]: Enhanced content and error message
        """
        logger.info(f"Starting enterprise LLM processing for {source_filename}")

        all_content_parts = []
        errors = []

        # 1. Process metadata first with metadata model
        logger.info("Processing metadata...")
        metadata_content = self._prepare_metadata_content(metadata, source_filename)
        enhanced_metadata, metadata_error = self.client.call_model(
            METADATA_PROCESSING_PROMPT,
            metadata_content,
            model_type="metadata"
        )

        if metadata_error:
            errors.append(f"Metadata processing error: {metadata_error}")

        all_content_parts.append("# Document Analysis\n\n" + enhanced_metadata)

        # 2. Process slides in batches of 5 with content model
        logger.info("Processing slide content in batches...")
        slides = structured_data.get("slides", [])

        # Separate regular slides from diagram slides
        regular_slides = []
        diagram_slides = []

        for slide in slides:
            if self._contains_diagrams(slide):
                diagram_slides.append(slide)
            else:
                regular_slides.append(slide)

        # Process regular slides in batches of 5
        enhanced_slides = []
        for i in range(0, len(regular_slides), 5):
            batch = regular_slides[i:i + 5]
            batch_content = self._prepare_slide_batch_content(batch, i + 1)

            enhanced_batch, batch_error = self.client.call_model(
                CONTENT_PROCESSING_PROMPT,
                batch_content,
                model_type="content"
            )

            if batch_error:
                errors.append(f"Batch {i // 5 + 1} processing error: {batch_error}")
                enhanced_slides.append(batch_content)  # Use original on error
            else:
                enhanced_slides.append(enhanced_batch)

        # 3. Process diagram slides with diagram model
        logger.info("Processing diagram content...")
        for slide in diagram_slides:
            slide_content = self._prepare_single_slide_content(slide)

            enhanced_slide, slide_error = self.client.call_model(
                DIAGRAM_PROCESSING_PROMPT,
                slide_content,
                model_type="diagram"
            )

            if slide_error:
                errors.append(f"Diagram slide {slide['slide_number']} error: {slide_error}")
                enhanced_slides.append(slide_content)  # Use original on error
            else:
                enhanced_slides.append(enhanced_slide)

        # 4. Combine all content parts
        all_content_parts.extend(enhanced_slides)

        # 5. Create final combined content
        final_content = "\n\n---\n\n".join(all_content_parts)

        # Add processing summary
        final_content += f"\n\n---\n\n## Processing Summary\n\n"
        final_content += f"- **Total slides processed:** {len(slides)}\n"
        final_content += f"- **Regular content batches:** {(len(regular_slides) + 4) // 5}\n"
        final_content += f"- **Diagram slides:** {len(diagram_slides)}\n"
        final_content += f"- **Models used:** Metadata, Content, Diagram\n"

        if errors:
            final_content += f"- **Processing errors:** {len(errors)}\n"
            for error in errors:
                final_content += f"  - {error}\n"

        error_message = "; ".join(errors) if errors else None

        logger.info(f"Enterprise LLM processing completed for {source_filename}")
        return final_content, error_message

    def _prepare_metadata_content(self, metadata: Dict, filename: str) -> str:
        """Prepare metadata for processing"""
        content = f"**File:** {filename}\n\n"
        content += "**PowerPoint Metadata:**\n\n"

        for key, value in metadata.items():
            if value:
                content += f"- **{key.replace('_', ' ').title()}:** {value}\n"

        content += "\n\nPlease analyse this metadata and provide insights for document management."
        return content

    def _prepare_slide_batch_content(self, slides: List[Dict], batch_number: int) -> str:
        """Prepare a batch of slides for processing"""
        content = f"**Slide Batch {batch_number}** (Slides {slides[0]['slide_number']}-{slides[-1]['slide_number']})\n\n"

        for slide in slides:
            content += f"### Slide {slide['slide_number']}\n\n"
            content += self._extract_slide_text_content(slide)
            content += "\n\n"

        content += "Please clean up and enhance this slide content whilst preserving all information."
        return content

    def _prepare_single_slide_content(self, slide: Dict) -> str:
        """Prepare a single slide for diagram processing"""
        content = f"**Slide {slide['slide_number']} (Contains Diagrams/Charts)**\n\n"
        content += self._extract_slide_text_content(slide)
        content += "\n\nPlease enhance this visual content and suggest diagram representations."
        return content

    def _extract_slide_text_content(self, slide: Dict) -> str:
        """Extract text content from a slide"""
        content_parts = []

        for block in slide.get("content_blocks", []):
            if block.get("type") == "text":
                for para in block.get("paragraphs", []):
                    clean_text = para.get("clean_text", "")
                    if clean_text:
                        hints = para.get("hints", {})
                        if hints.get("is_bullet"):
                            level = hints.get("bullet_level", 0)
                            indent = "  " * level
                            content_parts.append(f"{indent}- {clean_text}")
                        else:
                            content_parts.append(clean_text)

            elif block.get("type") == "table":
                content_parts.append("[TABLE CONTENT]")
                # Add basic table representation

            elif block.get("type") == "chart":
                content_parts.append(f"[CHART: {block.get('title', 'Untitled')}]")

            elif block.get("type") == "image":
                content_parts.append(f"[IMAGE: {block.get('alt_text', 'Image')}]")

        return "\n".join(content_parts)

    def _contains_diagrams(self, slide: Dict) -> bool:
        """Check if a slide contains diagrams or charts"""
        for block in slide.get("content_blocks", []):
            if block.get("type") in ["chart", "diagram"]:
                return True
            # Check for high concentration of shapes/lines (diagram indicators)
            if block.get("type") == "shape" or "arrow" in str(block) or "line" in str(block):
                return True
        return False


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

        # Return basic formatted content as fallback
        fallback_content = f"# {source_filename}\n\n"
        fallback_content += "**Error:** Enterprise LLM processing failed, showing original content.\n\n"

        # Add basic slide content
        for slide in structured_data.get("slides", []):
            fallback_content += f"## Slide {slide['slide_number']}\n\n"
            fallback_content += enhancer._extract_slide_text_content(slide) if 'enhancer' in locals() else str(slide)
            fallback_content += "\n\n"

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
