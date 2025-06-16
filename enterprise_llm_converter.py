"""
Simple Enterprise LLM Converter - Debug Version
No fallbacks, just direct connection testing
"""

import os
import json
import requests
import logging
from typing import Tuple, Optional, Dict

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Simple prompt from working Claude system
SIMPLE_PROMPT = """
You are a PowerPoint to Markdown converter. Clean up this markdown:

1. Fix bullet point hierarchies - use 2-space indentation
2. Make slide titles into # headers
3. Preserve ALL content and hyperlinks
4. Use proper markdown syntax

Keep ALL original text content. Output clean, readable markdown.
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
        Simple model call with detailed logging
        """
        logger.info("🚀 Calling enterprise model...")

        # Simple payload
        payload = {
            "messages": [
                {"role": "system", "content": SIMPLE_PROMPT},
                {"role": "user", "content": content}
            ],
            "max_tokens": 2000,
            "temperature": 0.1
        }

        logger.info(f"📤 Payload size: {len(json.dumps(payload))} characters")
        logger.info(f"📤 Headers: {list(self.headers.keys())}")

        try:
            response = requests.post(
                self.model_url,
                headers=self.headers,
                json=payload,
                timeout=60
            )

            logger.info(f"📥 Response status: {response.status_code}")
            logger.info(f"📥 Response headers: {dict(response.headers)}")

            if response.status_code == 200:
                result = response.json()
                logger.info(f"📥 Response keys: {list(result.keys())}")

                # Try different response formats
                if "choices" in result and result["choices"]:
                    content = result["choices"][0]["message"]["content"]
                    logger.info("✅ Extracted content from choices format")
                elif "generated_text" in result:
                    content = result["generated_text"]
                    logger.info("✅ Extracted content from generated_text format")
                elif "content" in result:
                    content = result["content"]
                    logger.info("✅ Extracted content from content format")
                else:
                    content = str(result)
                    logger.warning("⚠️ Using raw response as content")

                logger.info(f"✅ Success! Generated {len(content)} characters")
                return content, None

            else:
                error_msg = f"API error {response.status_code}: {response.text}"
                logger.error(f"❌ {error_msg}")
                return "", error_msg

        except requests.exceptions.Timeout:
            error_msg = "Request timeout (60 seconds)"
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
        Simple processing - just test the connection
        """
        logger.info(f"🎯 Processing {source_filename}...")

        # Create simple test content
        test_content = f"# Test Document: {source_filename}\n\n"
        test_content += "- First bullet point\n"
        test_content += "- Second bullet point\n"
        test_content += "\nThis is a test of the enterprise LLM connection."

        logger.info(f"📝 Test content: {len(test_content)} characters")

        # Call the model
        enhanced_content, error = self.client.call_model(test_content)

        if error:
            logger.error(f"❌ Enhancement failed: {error}")
            return test_content, error

        logger.info("✅ Enhancement successful!")
        return enhanced_content, None


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