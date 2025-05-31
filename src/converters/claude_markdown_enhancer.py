"""
Claude Sonnet 4 Markdown Enhancement Module

This module provides functions to enhance markdown documents using Claude Sonnet 4
with the document-to-markdown conversion system prompt.
"""

import os
import anthropic
from typing import Tuple, Optional
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# The system prompt for document to markdown conversion
DOCUMENT_TO_MARKDOWN_SYSTEM_PROMPT = """
# Document Conversion System Prompt

## Role
You are a document conversion specialist. Convert PowerPoint presentations, Word documents, and PDFs into standardized Markdown format that preserves 100% of original content while optimizing for vector database storage.

## Critical Rules
- **NEVER remove, summarize, or paraphrase any original content**
- **NEVER skip or omit any text, images, or formatting elements**
- **ALWAYS preserve exact wording, including typos and inconsistencies**
- **ALWAYS maintain original structure and reading order**

---

## Document Header Format (Required)

```markdown
# [Exact Document Title]
**Source:** [filename.extension]
**Content Type:** [PowerPoint/Word/PDF]
**Conversion Date:** [YYYY-MM-DD]
**Total Slides/Pages:** [count]
**Document ID:** [source-filename-timestamp]
---
```

## Slide/Section Structure (Required)

```markdown
## Slide [Number]: [Exact Title]

### [Content Type]: [Brief Description]
[Preserved content exactly as written]

### [Additional Content Type]: [Description]
[Additional content if present]

---
```

## Content Type Labels (Use These Exactly)
- **Main Content:** Primary text and bullet points
- **Textbox Content:** Standalone text elements
- **Shape Content:** Text from shapes and objects
- **Speaker Notes:** Presenter notes
- **Image Content:** Captions and alt-text
- **Table Content:** Structured data
- **Footer/Header:** Document metadata

## Formatting Rules

### Headings
- `#` Document title ONLY
- `##` Slide titles (format: "Slide X: Title")
- `###` Content type labels
- `####` Sub-sections (rare use)

### Lists
- **Indentation:** 2 spaces per level
- **Preserve nesting:** Maintain original depth exactly
- **Fix obvious breaks:** Repair clear formatting errors while preserving all text

### Text Formatting
- **Bold:** `**text**`
- **Italic:** `*text*`
- **Code:** `backticks`
- **Links:** `[text](URL)` - never modify URLs
- **Preserve:** All Unicode, emojis, special characters

## Content Processing

### Reading Order
Arrange content in logical reading sequence when original order is unclear.

### Duplicate Content
If identical content appears multiple times, consolidate but note in output.

### Unclear Content
Preserve exactly as written. Use `[unclear: original-text]` only if critical for understanding.

### Missing Context
Never add explanatory text. Preserve exactly what was provided.

## Quality Standards

Each conversion must include:
- [ ] Complete document header with all required fields
- [ ] Consistent slide structure throughout
- [ ] All content type labels applied correctly
- [ ] All original text preserved without changes
- [ ] List formatting standardized with proper nesting
- [ ] All links preserved with original URLs
- [ ] Special characters and Unicode intact
- [ ] Logical reading order maintained
- [ ] Valid markdown syntax throughout

## Error Handling

**If content is unclear:** Preserve exactly as written
**If structure is ambiguous:** Use most literal interpretation
**If formatting is broken:** Fix formatting while preserving all text
**If elements are missing:** Note what's missing, never invent content

## Output Template

```markdown
# [Document Title]
**Source:** [filename]
**Content Type:** [type]
**Conversion Date:** [date]
**Total Slides/Pages:** [count]
**Document ID:** [unique-id]
---

## Slide 1: [Title]
### Main Content: [Description]
[Content]
---

## Slide 2: [Title]
### [Content Type]: [Description]
[Content]
---
```

Remember: Your job is to convert documents with perfect accuracy and consistent formatting. Focus only on the individual document you're processing.
"""



class ClaudeMarkdownEnhancer:
    """
    A class to enhance markdown documents using Claude Sonnet 4.
    """

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize the Claude Markdown Enhancer.

        Args:
            api_key (Optional[str]): Anthropic API key. If not provided, will try to get from environment.
        """
        self.api_key = api_key or os.getenv("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Anthropic API key is required. Set ANTHROPIC_API_KEY environment variable or pass it directly.")

        self.client = anthropic.Anthropic(api_key=self.api_key)
        self.model = "claude-sonnet-4-20250514"

    def enhance_markdown(self, markdown_content: str, source_filename: str = "unknown",
                         content_type: str = "Document") -> Tuple[str, Optional[str]]:
        """
        Enhance markdown content using Claude Sonnet 4.

        Args:
            markdown_content (str): The original markdown content to enhance
            source_filename (str): The source filename for context
            content_type (str): The type of document (PowerPoint, Word, PDF, etc.)

        Returns:
            Tuple[str, Optional[str]]: Enhanced markdown content and error message (if any)
        """
        try:
            # Create the user prompt with context
            user_prompt = f"""Please enhance this markdown document according to the document-to-markdown conversion system prompt. 

**Source Information:**
- Filename: {source_filename}
- Content Type: {content_type}

**Original Markdown Content:**
{markdown_content}

Please apply the formatting standards and ensure all content is preserved while improving the structure, consistency, and readability. Make sure to include proper document headers and maintain all original content exactly as provided."""

            # Make the API call
            response = self.client.messages.create(
                model=self.model,
                max_tokens=4096,
                temperature=0.1,  # Low temperature for consistent formatting
                system=DOCUMENT_TO_MARKDOWN_SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": user_prompt
                    }
                ]
            )

            enhanced_content = response.content[0].text
            logger.info(f"Successfully enhanced markdown for {source_filename}")
            return enhanced_content, None

        except anthropic.APIError as e:
            error_msg = f"Anthropic API error: {str(e)}"
            logger.error(error_msg)
            return markdown_content, error_msg

        except Exception as e:
            error_msg = f"Unexpected error enhancing markdown: {str(e)}"
            logger.error(error_msg)
            return markdown_content, error_msg

    def enhance_multiple_documents(self, documents: list) -> list:
        """
        Enhance multiple markdown documents.

        Args:
            documents (list): List of dictionaries with keys 'content', 'filename', 'content_type'

        Returns:
            list: List of results with enhanced content and any errors
        """
        results = []

        for doc in documents:
            content = doc.get('content', '')
            filename = doc.get('filename', 'unknown')
            content_type = doc.get('content_type', 'Document')

            enhanced_content, error = self.enhance_markdown(content, filename, content_type)

            results.append({
                'filename': filename,
                'original_content': content,
                'enhanced_content': enhanced_content,
                'error': error,
                'success': error is None
            })

        return results


def enhance_markdown_with_claude(markdown_content: str, api_key: str,
                                 source_filename: str = "unknown",
                                 content_type: str = "Document") -> Tuple[str, Optional[str]]:
    """
    Convenience function to enhance markdown content using Claude Sonnet 4.

    Args:
        markdown_content (str): The markdown content to enhance
        api_key (str): Anthropic API key
        source_filename (str): Source filename for context
        content_type (str): Document type

    Returns:
        Tuple[str, Optional[str]]: Enhanced content and error message (if any)
    """
    try:
        enhancer = ClaudeMarkdownEnhancer(api_key)
        return enhancer.enhance_markdown(markdown_content, source_filename, content_type)
    except Exception as e:
        return markdown_content, str(e)


# Example usage and testing
if __name__ == "__main__":
    # Example markdown content
    sample_markdown = """
# My Document

This is some content with poor formatting.

* Item 1
* Item 2
  * Nested item
* Item 3

Some text here.

## Section 2
More content here with [a link](https://example.com).

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
"""

    # Example usage
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if api_key:
        enhancer = ClaudeMarkdownEnhancer(api_key)
        enhanced, error = enhancer.enhance_markdown(
            sample_markdown,
            "sample_document.md",
            "Markdown Document"
        )

        if error:
            print(f"Error: {error}")
        else:
            print("Enhanced Markdown:")
            print("=" * 50)
            print(enhanced)
    else:
        print("Please set ANTHROPIC_API_KEY environment variable to test.")