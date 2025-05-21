"""
Office to Markdown - Configuration Settings

This file contains all configuration settings for the application.
"""

# App Settings
APP_TITLE = "Office to Markdown"
APP_ICON = "📄"
APP_LAYOUT = "centered"

# UI Settings
UI_THEME_COLOR = "#4e54c8"
UI_BACKGROUND_COLOR = "#f5f7f9"

# File Settings
DEFAULT_MARKDOWN_SUBFOLDER = "markdown"

# API Settings
OPENAI_MODEL = "gpt-4o"
OPENAI_TEMPERATURE = 0.3
OPENAI_MAX_TOKENS = 8000

# Hyperlink Extraction Settings
HYPERLINK_CONTEXT_SIZE = 10  # Words to extract around a hyperlink for context

# File Format Categories
FILE_FORMATS = {
    "📝 Documents": {
        "formats": ["Word (.docx, .doc)", "PDF (with hyperlink extraction)", "EPub"],
        "extensions": ["docx", "doc", "pdf", "epub"],
    },
    "📊 Spreadsheets": {
        "formats": ["Excel (.xlsx, .xls)"],
        "extensions": ["xlsx", "xls"],
    },
    "📊 Presentations": {
        "formats": ["PowerPoint (.pptx, .ppt) with hyperlink extraction"],
        "extensions": ["pptx", "ppt"],
    },
    "🌐 Web": {"formats": ["HTML"], "extensions": ["html", "htm"]},
    "📁 Others": {
        "formats": ["CSV", "JSON", "XML", "ZIP (iterates over contents)"],
        "extensions": ["csv", "json", "xml", "zip"],
    },
}

# User Agent for web requests
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"

# AI Enhancement prompts
MARKDOWN_ENHANCEMENT_PROMPT = """You are a markdown formatting expert. Your task is to:
1. Fix any syntax errors in the markdown
2. Improve the structure and hierarchy of headers
3. Ensure consistent formatting throughout
4. Enhance bullet points and numbered lists
5. Properly format tables and code blocks
6. Add appropriate spacing between sections
7. Maintain the original content without adding new information
8. Preserve all links and references

Return ONLY the enhanced markdown content without any explanations or additional text."""