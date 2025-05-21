
from markitdown import MarkItDown

# Create the converter instance
converter = MarkItDown()

# Convert the PDF file to markdown
result = converter.convert("pdf_test_document.pdf")

# Write the markdown content to a file
# The result object has a 'markdown' attribute (not 'content')
with open("output.md", "w") as f:
    f.write(result.markdown)