# Office to Markdown Converter 📄

A powerful Streamlit web application that converts various document formats and websites to clean, structured Markdown with advanced formatting preservation and hyperlink extraction.

## ✨ Features

### Document Conversion
- **📝 Documents**: Word (.docx, .doc), PDF (with hyperlink extraction), EPub
- **📊 Spreadsheets**: Excel (.xlsx, .xls)
- **📊 Presentations**: PowerPoint (.pptx, .ppt) with enhanced formatting and hyperlink extraction
- **🌐 Web Content**: HTML files and live website URLs
- **📁 Other Formats**: CSV, JSON, XML, ZIP archives

### Advanced PowerPoint Processing
- **Enhanced Formatting Preservation**: Bold, italic, underline, and font styling
- **Smart Hierarchy Detection**: Automatic header levels and proper nesting
- **Bullet Point & List Support**: Multi-level bullet points and numbered lists
- **Table Conversion**: Full table support with formatting
- **Inline Hyperlink Extraction**: Preserves clickable links within text
- **Reading Order Optimization**: Processes content in logical reading sequence

### Hyperlink Extraction
- **PDF Hyperlinks**: Extracts both external URLs and internal document links
- **PowerPoint Hyperlinks**: Captures shape-level and text-level hyperlinks
- **Context Preservation**: Maintains link context and descriptive text
- **Duplicate Handling**: Smart deduplication with text quality prioritization

### AI Enhancement
- **OpenAI Integration**: Optional AI-powered markdown formatting improvement
- **Syntax Correction**: Fixes markdown syntax errors
- **Structure Optimization**: Improves header hierarchy and formatting consistency
- **Content Preservation**: Enhances formatting without altering original content

### Batch Processing
- **Folder Processing**: Convert entire directories of documents
- **Progress Tracking**: Real-time conversion progress with detailed status
- **Error Reporting**: Comprehensive error handling and reporting
- **Flexible Output**: Customizable output directory structure

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Installation

1. **Clone the repository:**
```bash
git clone https://github.com/yourusername/office-to-markdown.git
cd office-to-markdown
```

2. **Install dependencies:**
```bash
pip install -r requirements.txt
```

3. **Run the application:**
```bash
streamlit run app.py
```

4. **Open your browser** and navigate to `http://localhost:8501`

## 🔧 Configuration

### Environment Variables
Create a `.env` file in the root directory (optional):
```env
OPENAI_API_KEY=your_openai_api_key_here
```

### Config Settings
Modify `config.py` to customize:
- **UI Theme**: Colors and styling
- **AI Model**: OpenAI model selection and parameters
- **File Formats**: Add or modify supported formats
- **Enhancement Prompts**: Customize AI enhancement behavior

## 📖 Usage

### Single File Conversion

1. **Select the "File Upload" tab**
2. **Choose your document** from the supported formats
3. **Configure options**:
   - Enable/disable AI enhancement
   - Provide OpenAI API key (for enhancement)
4. **Click "Convert File to Markdown"**
5. **Download the result** or copy from the text area

### Website Conversion

1. **Select the "Website URL" tab**
2. **Enter the complete URL** (including http:// or https://)
3. **Configure enhancement options**
4. **Click "Convert Website to Markdown"**
5. **Download the converted content**

### Batch Folder Processing

1. **Select the "Folder Processing" tab**
2. **Enter the input folder path** containing your documents
3. **Optionally specify output folder** (defaults to "markdown" subfolder)
4. **Configure enhancement settings**
5. **Click "Process Folder"**
6. **Monitor progress** and view results summary

## 🏗️ Architecture

### Project Structure
```
office-to-markdown/
├── app.py                          # Main Streamlit application
├── config.py                       # Configuration settings
├── requirements.txt                 # Python dependencies
├── src/
│   ├── converters/                 # Document conversion modules
│   │   ├── file_converter.py       # Main file conversion logic
│   │   ├── url_converter.py        # Website conversion
│   │   ├── enhanced_pptx_processor.py  # Enhanced PowerPoint processing
│   │   └── hyperlink_extractor.py  # Hyperlink extraction utilities
│   ├── processors/                 # Batch processing
│   │   └── folder_processor.py     # Folder batch operations
│   ├── ui/                        # User interface components
│   │   └── components.py          # Streamlit UI helpers
│   └── utils/                     # Utility functions
│       └── file_utils.py          # File handling utilities
└── tests/                         # Unit tests
    ├── test_file_converter.py
    ├── test_url_converter.py
    ├── test_hyperlink_extractor.py
    └── test_file_utils.py
```

### Key Components

#### Enhanced PowerPoint Processor
- **Comprehensive Shape Processing**: Handles text boxes, placeholders, tables, groups, and images
- **Format Preservation**: Maintains bold, italic, underline, and font styling
- **Smart List Detection**: Recognizes and properly formats bullet points and numbered lists
- **Hyperlink Integration**: Preserves both inline and shape-level hyperlinks
- **Reading Order**: Processes content based on spatial positioning

#### Hyperlink Extraction System
- **Multi-format Support**: Works with PDF and PowerPoint documents
- **Context Analysis**: Extracts surrounding text for meaningful link descriptions
- **Duplicate Management**: Intelligent handling of repeated URLs with quality-based text selection
- **Structured Output**: Organizes links by page/slide with proper markdown formatting

## 🔧 API Reference

### File Converter
```python
from src.converters.file_converter import convert_file_to_markdown

# Convert any supported file
markdown_content, error = convert_file_to_markdown(
    file_data=file_bytes,
    filename="document.docx",
    enhance=True,
    api_key="your_openai_key"
)
```

### URL Converter
```python
from src.converters.url_converter import convert_url_to_markdown

# Convert website to markdown
markdown_content, error, title = convert_url_to_markdown(
    url="https://example.com",
    enhance=True,
    api_key="your_openai_key"
)
```

### Hyperlink Extraction
```python
from src.converters.hyperlink_extractor import extract_pdf_hyperlinks, extract_pptx_hyperlinks

# Extract hyperlinks from documents
pdf_links = extract_pdf_hyperlinks("document.pdf")
pptx_links = extract_pptx_hyperlinks("presentation.pptx")
```

## 🧪 Testing

Run the test suite:
```bash
python -m pytest tests/ -v
```

Run specific test modules:
```bash
python -m pytest tests/test_file_converter.py -v
python -m pytest tests/test_hyperlink_extractor.py -v
```

## 📋 Dependencies

### Core Libraries
- **streamlit**: Web application framework
- **markitdown**: Primary document conversion engine
- **python-pptx**: Enhanced PowerPoint processing
- **PyMuPDF**: PDF hyperlink extraction
- **beautifulsoup4**: HTML processing
- **openai**: AI enhancement capabilities

### Full Dependency List
See `requirements.txt` for complete list of dependencies with version specifications.

## 🤝 Contributing

We welcome contributions! Please follow these steps:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/amazing-feature`
3. **Make your changes** and add tests
4. **Run the test suite**: `python -m pytest tests/ -v`
5. **Commit your changes**: `git commit -m 'Add amazing feature'`
6. **Push to the branch**: `git push origin feature/amazing-feature`
7. **Open a Pull Request**

### Development Guidelines
- Follow PEP 8 style guidelines
- Add unit tests for new functionality
- Update documentation for API changes
- Ensure all tests pass before submitting

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Known Issues & Limitations

- **PDF Processing**: Hyperlink extraction works best with well-structured PDFs; consider using original document formats when possible
- **Complex Layouts**: Very complex PowerPoint layouts may require manual review
- **Large Files**: Processing very large files may take significant time
- **API Rate Limits**: OpenAI enhancement is subject to API rate limiting

## 🚀 Roadmap

- [ ] Support for additional document formats (OneNote, Google Docs)
- [ ] Advanced table formatting options
- [ ] Image extraction and embedding
- [ ] Custom enhancement templates
- [ ] API endpoint for programmatic access
- [ ] Docker containerization
- [ ] Cloud deployment options

## 💡 Tips for Best Results

1. **Use Original Formats**: When possible, use Word documents instead of PDFs for better formatting preservation
2. **Check Hyperlinks**: Review extracted hyperlinks for accuracy, especially in complex documents
3. **AI Enhancement**: Experiment with AI enhancement to improve markdown structure
4. **Batch Processing**: Use folder processing for converting multiple documents efficiently
5. **Quality Review**: Always review converted content for formatting accuracy

## 📞 Support

If you encounter any issues or have questions:

1. **Check the Issues**: Search existing GitHub issues
2. **Create an Issue**: Open a new issue with detailed information
3. **Documentation**: Review this README and code comments
4. **Tests**: Run the test suite to verify your environment

## 🙏 Acknowledgments

- **MarkItDown**: Core conversion engine
- **python-pptx**: PowerPoint processing capabilities
- **PyMuPDF**: PDF processing and hyperlink extraction
- **Streamlit**: Excellent web framework for Python applications
- **OpenAI**: AI enhancement capabilities

---

**Developed by James Taylor**

Made with ❤️ for the document conversion community