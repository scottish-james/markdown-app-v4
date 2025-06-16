"""
Simple Enterprise LLM Converter - Debug Version
No fallbacks, just direct connection testing
"""

import os
import json
import requests
import logging
import re
from typing import Tuple, Optional, Dict

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Proven prompt from working Claude system for slide batches
SLIDE_BATCH_PROMPT = """
You are a PowerPoint to Markdown converter. You receive a batch of PowerPoint slides (max 5 slides) and must clean them up into professional, well-structured markdown.

Your job:
1. Fix bullet point hierarchies - create proper nested lists with 2-space indentation
2. Identify slide titles and format as # headers
3. Identify sub headings and format as ### headers 
4. Preserve ALL hyperlinks and formatting (bold, italic)
5. Fix broken list structures
6. Ensure tables are properly formatted
7. Clean up spacing and structure
8. Keep slide markers like <!-- Slide 1 --> for reference

Key Rules:
- Keep ALL original text content - this is critical for regulatory documents
- Fix bullet nesting based on context and content
- Make short, standalone text into appropriate ### headers 
- Preserve all hyperlinks exactly as provided
- Use proper markdown syntax throughout
- Look for numbered sequences that are out of order and fix them

The input will have slide markers like <!-- Slide 1 --> to separate slides. Keep these markers in your output.

Output clean, readable markdown that maintains the original document's intent but fixes structure and formatting issues.
"""


class EnterpriseLLMClient:
    """
    Simple client for testing enterprise LLM connection
    """

    def __init__(self):
        """Initialize and test connection immediately"""
        logger.info("🔧 Initializing Enterprise LLM Client...")

        self.jwt_token = self._load_jwt_token()
        logger.info(f"✅ JWT token loaded: {self.jwt_token[:20]}...")

        self.model_url = self._load_model_url()
        logger.info(f"✅ Model URL loaded: {self.model_url}")

        self.headers = {
            "Authorization": f"Bearer {self.jwt_token}",
            "Content-Type": "application/json"
        }

        # Test connection immediately
        self._test_connection()

    def _load_jwt_token(self) -> str:
        """Load JWT token with detailed logging"""
        logger.info("📄 Loading JWT token...")

        if not os.path.exists("JWT_token.txt"):
            raise ValueError("❌ JWT_token.txt file not found")

        with open("JWT_token.txt", "r") as f:
            token = f.read().strip()

        if not token:
            raise ValueError("❌ JWT token file is empty")

        if not token.count('.') == 2:
            logger.warning(f"⚠️ JWT token format looks unusual: {token.count('.')} dots (expected 2)")

        return token

    def _load_model_url(self) -> str:
        """Load model URL with detailed logging"""
        logger.info("🌐 Loading model URL...")

        if not os.path.exists("model_url.txt"):
            raise ValueError("❌ model_url.txt file not found")

        with open("model_url.txt", "r") as f:
            content = f.read().strip()

        if not content:
            raise ValueError("❌ Model URL file is empty")

        # Handle both single URL and JSON format
        if content.startswith('{'):
            logger.info("📋 JSON format detected")
            urls = json.loads(content)
            # Use content model or first available
            url = urls.get("content", list(urls.values())[0])
            logger.info(f"📋 Using URL from JSON: {url}")
        else:
            logger.info("📋 Single URL format detected")
            url = content

        if not url.startswith(('http://', 'https://')):
            logger.warning(f"⚠️ URL doesn't start with http/https: {url}")

        return url

    def _test_connection(self):
        """Test basic connectivity to the endpoint"""
        logger.info("🧪 Testing connection to enterprise endpoint...")

        try:
            # Simple connectivity test (no auth)
            response = requests.head(self.model_url, timeout=10)
            logger.info(f"✅ Endpoint reachable (status: {response.status_code})")
        except requests.exceptions.ConnectTimeout:
            logger.error("❌ Connection timeout - endpoint unreachable")
            raise
        except requests.exceptions.ConnectionError as e:
            logger.error(f"❌ Connection failed: {e}")
            raise
        except Exception as e:
            logger.warning(f"⚠️ Connection test inconclusive: {e}")

    def call_model(self, content: str) -> Tuple[str, Optional[str]]:
        """
        Call model with slide batch prompt
        """
        logger.info("🚀 Calling enterprise model...")

        # Use slide batch prompt
        payload = {
            "messages": [
                {"role": "system", "content": SLIDE_BATCH_PROMPT},
                {"role": "user", "content": content}
            ],
            "max_tokens": 4000,  # Increased for batch processing
            "temperature": 0.1
        }

        logger.info(f"📤 Payload size: {len(json.dumps(payload))} characters")

        try:
            response = requests.post(
                self.model_url,
                headers=self.headers,
                json=payload,
                timeout=120  # Increased timeout for batches
            )

            logger.info(f"📥 Response status: {response.status_code}")

            if response.status_code == 200:
                result = response.json()

                # Try different response formats
                if "choices" in result and result["choices"]:
                    content = result["choices"][0]["message"]["content"]
                elif "generated_text" in result:
                    content = result["generated_text"]
                elif "content" in result:
                    content = result["content"]
                else:
                    content = str(result)

                logger.info(f"✅ Success! Generated {len(content)} characters")
                return content, None

            else:
                error_msg = f"API error {response.status_code}: {response.text}"
                logger.error(f"❌ {error_msg}")
                return "", error_msg

        except Exception as e:
            error_msg = f"Request failed: {str(e)}"
            logger.error(f"❌ {error_msg}")
            return "", error_msg


class EnterpriseLLMEnhancer:
    """
    Simple enhancer for testing
    """

    def __init__(self):
        """Initialize with immediate connection test"""
        logger.info("🎯 Initializing Enterprise LLM Enhancer...")
        self.client = EnterpriseLLMClient()
        logger.info("✅ Enterprise LLM Enhancer ready")

    def enhance_powerpoint_content(self, structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
    Tuple[str, Optional[str]]:
        """
        Process PowerPoint content in batches of 5 slides
        """
        logger.info(f"🎯 Processing {source_filename}...")

        # First convert structured data to basic markdown from PowerPoint processor
        basic_markdown = self._convert_structured_to_basic_markdown(structured_data, metadata, source_filename)
        logger.info(f"📝 Generated basic markdown: {len(basic_markdown)} characters")

        # Split into slides based on HTML markers
        slide_batches = self._split_into_slide_batches(basic_markdown)
        logger.info(f"📊 Split into {len(slide_batches)} batches of max 5 slides each")

        # Process metadata separately
        metadata_content = self._process_metadata(metadata, source_filename)

        # Process each batch
        enhanced_parts = [metadata_content] if metadata_content else []

        for i, batch in enumerate(slide_batches):
            logger.info(f"🚀 Processing batch {i + 1}/{len(slide_batches)}...")

            enhanced_batch, error = self.client.call_model(batch)

            if error:
                logger.error(f"❌ Batch {i + 1} failed: {error}")
                enhanced_parts.append(batch)  # Use original on error
            else:
                logger.info(f"✅ Batch {i + 1} enhanced successfully")
                enhanced_parts.append(enhanced_batch)

        # Combine all parts
        final_content = "\n\n---\n\n".join(enhanced_parts)

        logger.info(f"✅ Final content: {len(final_content)} characters")
        return final_content, None

    def _convert_structured_to_basic_markdown(self, structured_data: Dict, metadata: Dict, source_filename: str) -> str:
        """
        Convert structured PowerPoint data to basic markdown with slide markers
        """
        markdown_parts = []

        # Add metadata as HTML comment for the LLM
        if metadata:
            markdown_parts.append("<!-- POWERPOINT METADATA FOR CLAUDE:")
            for key, value in metadata.items():
                if value:
                    markdown_parts.append(f"{key}: {value}")
            markdown_parts.append("-->")

        # Process each slide with HTML markers
        for slide in structured_data.get("slides", []):
            slide_parts = []
            slide_parts.append(f"<!-- Slide {slide['slide_number']} -->")

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
                                slide_parts.append(f"{indent}- {text}")
                            else:
                                slide_parts.append(text)

                elif block.get("type") == "table":
                    table_data = block.get("data", [])
                    if table_data:
                        for i, row in enumerate(table_data):
                            slide_parts.append("| " + " | ".join(row) + " |")
                            if i == 0:  # Header separator
                                slide_parts.append("| " + " | ".join("---" for _ in row) + " |")

                elif block.get("type") == "image":
                    alt_text = block.get("alt_text", "Image")
                    slide_parts.append(f"![{alt_text}](image)")

                elif block.get("type") == "chart":
                    title = block.get("title", "Chart")
                    slide_parts.append(f"**Chart: {title}**")
                    # Ignore diagram markers for now as requested

            # Add slide content if not empty
            if len(slide_parts) > 1:  # More than just the slide marker
                markdown_parts.append("\n".join(slide_parts))

        return "\n\n".join(markdown_parts)

    def _split_into_slide_batches(self, markdown_content: str) -> list:
        """
        Split markdown content into batches of 5 slides based on HTML markers
        """
        import re

        # Split by slide markers
        slide_pattern = r'<!-- Slide (\d+) -->'
        slides = re.split(slide_pattern, markdown_content)

        # Remove empty parts and metadata
        clean_slides = []
        current_slide = ""

        for i, part in enumerate(slides):
            if re.match(r'^\d+


def enhance_markdown_with_enterprise_llm(structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
Tuple[str, Optional[str]]:
    """
    Simple test function
    """
    try:
        enhancer = EnterpriseLLMEnhancer()
        return enhancer.enhance_powerpoint_content(structured_data, metadata, source_filename)
    except Exception as e:
        error_msg = f"Enterprise LLM failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)  # Don't fall back - let it fail so we can debug

, part):  # This is a slide number
if current_slide:
    clean_slides.append(current_slide.strip())
current_slide = f"<!-- Slide {part} -->"
elif part.strip() and not part.startswith("<!-- POWERPOINT METADATA"):
current_slide += "\n" + part

# Add final slide
if current_slide:
    clean_slides.append(current_slide.strip())

# Group into batches of 5
batches = []
for i in range(0, len(clean_slides), 5):
    batch = clean_slides[i:i + 5]
batches.append("\n\n".join(batch))

return batches


def _process_metadata(self, metadata: Dict, source_filename: str) -> str:
    """
    Process metadata separately
    """
    if not metadata:
        return ""

    metadata_parts = [f"# Document Analysis: {source_filename}", ""]

    for key, value in metadata.items():
        if value:
            clean_key = key.replace('_', ' ').title()
            metadata_parts.append(f"**{clean_key}:** {value}")

    return "\n".join(metadata_parts)


def enhance_markdown_with_enterprise_llm(structured_data: Dict, metadata: Dict, source_filename: str = "unknown") -> \
Tuple[str, Optional[str]]:
    """
    Simple test function
    """
    try:
        enhancer = EnterpriseLLMEnhancer()
        return enhancer.enhance_powerpoint_content(structured_data, metadata, source_filename)
    except Exception as e:
        error_msg = f"Enterprise LLM failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)  # Don't fall back - let it fail so we can debug