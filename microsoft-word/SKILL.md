---
name: microsoft-word
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Acts as a Microsoft Word wrapper around the docx skill. Triggers include: any mention of 'Word', 'Microsoft Word', '.docx file', requests for professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a 'report', 'memo', 'letter', 'template', or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
version: 1.0.0
license: Proprietary - See parent docx skill LICENSE.txt
metadata:
  hermes:
    tags: [documents, productivity, word-processing, microsoft-word, docx]
    related_skills: [docx, google-workspace, ocr-and-documents, openai-gateway-filesystem]
---

# Microsoft Word Skill - Office Document Operations

## Overview

The `microsoft-word` skill provides a high-level API for creating and manipulating Microsoft Word documents. It wraps the official Anthropic docx skill to provide more intuitive commands.

Think of this as a "virtual Microsoft Word" running in your terminal.

---

## Quick Start

```python
# Create a new document
from hermes_skills import microsoft_word

doc = microsoft_word.new_document(
    path='/home/user/report.docx',
    title="My Report"
)

# Add content
doc.add_heading("Executive Summary")
doc.add_paragraph("This is the summary...")

# Save
doc.save()
```

---

## Core Commands

### Create New Document

```python
from hermes_skills import microsoft_word

# Basic document
doc = microsoft_word.new_document(
    path='/home/user/mydocument.docx'
)
doc.title = "Untitled"
```

**With formatting:**

```python
doc = microsoft_word.new_document(
    path='/home/user/report.docx',
    title="Quarterly Report",
    author="Your Name",
    created=datetime.now(),
    page_size="Letter",  # or "A4"
    orientation="portrait"
)
```

### Reading Documents

**Read full content:**

```python
doc = microsoft_word.read_document(path='/path/to/file.docx')
print(doc.title)
print(doc.text_content)  # All text as plain string
print(doc.paragraphs)     # List of paragraph objects
```

**Get specific sections:**

```python
for para in doc.paragraphs:
    if para.style == 'Heading1':
        print(f"Section: {para.text}")
```

### Adding Content

**Headings:**

```python
doc.add_heading("Main Heading", level=1)  # Level 0-9
# or with explicit style
doc.add_heading("Main Heading", style="Heading1")
```

**Paragraphs:**

```python
doc.add_paragraph("Some text here...")

# Or add multiple lines
doc.add_paragraph(
    "Line 1.\n"
    "Line 2.\n"
    "Line 3."
)
```

### Tables

**Create a table:**

```python
table = doc.add_table(
    rows=3, columns=4,  # 3 rows x 4 columns
    headers=["Name", "Age", "City", "Email"]
)

# Add data
for i in range(3):
    for j in range(4):
        table[i][j].text = f"Data {i},{j}"
```

**Style a table:**

```python
table.style = "Table Grid"  # or "Table Classic", etc.
table.width = "90%"  # or DXA pixels like 8640
```

### Images

**Add image from file:**

```python
doc.add_image(
    path='/home/user/logo.png',
    caption="Company Logo",
    position="left"
)
```

**Add base64 image:**

```python
import base64
with open('chart.png', 'rb') as f:
    img_data = base64.b64encode(f.read()).decode()

doc.add_image(
    image_data=img_data,
    caption="Revenue Chart",
    style="Normal"
)
```

### Styles

**Override default styles:**

```python
doc.styles[0] = {  # Default paragraph style
    'font': {'name': 'Arial', 'size': 12},
    'bold': False,
    'italic': False
}

doc.styles['Heading1'] = {
    'font': {'name': 'Arial Black', 'size': 24},
    'bold': True,
    'color': '#2E75B6'
}
```

### Tracked Changes

**Add edit history:**

```python
from docx import Document as DWX
import datetime

doc = microsoft_word.read_document(path)
# Access redlining information
redlines = doc.redlines
for rl in redlines:
    print(f"{rl.author} modified at {rl.date}")
```

---

## Advanced Features

### Merge Documents

```python
from hermes_skills import microsoft_word

# Combine multiple documents
doc1 = microsoft_word.read_document('/first.docx')
doc2 = microsoft_word.read_document('/second.docx')

merged = microsoft_word.new_document()
merged.content.append(doc1)
merged.content.append(doc2)
merged.save('/combined.docx')
```

### Extract Tables

```python
doc = microsoft_word.read_document(path)

for table in doc.tables:
    print("Table found:")
    for row in table.rows:
        cells = [cell.text for cell in row.cells]
        print(cells)
```

### Find & Replace

**Simple replace:**

```python
doc.find_and_replace(
    old_text="Q3",
    new_text="Third Quarter"
)
```

**With formatting:**

```python
# Replace all instances of "important" with "IMPORTANT" in red text
doc.find_and_replace(
    search="important",
    replace="IMPORTANT",
    style="Heading1"  # Find only in Heading1 style paragraphs
)
```

---

## Command-Line Interface

### Create Document from Prompt

```bash
# Basic usage
microsoft-word create --path /home/user/report.docx --title "My Report"

# With options
microsoft-word create \
  --path /tmp/draft.docx \
  --author "Jane Smith" \
  --page-size A4
```

### Read and Extract

```bash
# Plain text extraction
microsoft-word extract /path/to/docx/file.docx

# As markdown (if pandoc available)
pandoc /path/to/docx/file.docx -o output.md
```

---

## Common Patterns

### Creating a Report Template

```python
doc = microsoft_word.new_document(
    path='/templates/quarterly_report.docx',
    title="Quarterly Business Report"
)

# Header with company info
from docx import Document as DWX
header_table = doc.add_table(rows=1, cols=3)
header_table.style = 'Light Shading Accent 1'

doc.add_heading("Executive Summary", level=0)
doc.add_paragraph("""This report covers the performance metrics...""")

# Table of contents (requires pandoc for proper extraction)
if microsoft_word.check_pandoc():
    doc.add_table_of_contents()

doc.save()
```

### Processing Batch Documents

```python
import glob
from hermes_skills import microsoft_word

files = sorted(glob.glob('/documents/*.docx'))
summary = {
    'total': len(files),
    'authors': {},
    'sizes': []
}

for f in files:
    doc = microsoft_word.read_document(f)
    summary['authors'][f] = doc.metadata.get('author', 'Unknown')
    summary['sizes'].append(os.path.getsize(f))

print(f"Processed {summary['total']} documents")
```

### Creating an Invoice Template

```python
from datetime import date

doc = microsoft_word.new_document(
    path='/invoices/invoice.docx',
    title="Invoice"
)

# Header
doc.add_heading("INVOICE", level=0)
doc.add_paragraph(f"Date: {date.today()}")

# Bill to section
section = doc.add_section("Bill To:")
section.text = """Acme Corporation\n
123 Business Street\n
New York, NY 10001"""

# Line items table
table = doc.add_table(rows=1, cols=5, headers=[
    "Description", "Quantity", "Unit Price", "Discount", "Total"
])
table.style = "Table Grid"

# Add product line
row = table[0]
row[0].text = "Consulting Services"
row[1].text = "10 hours"
row[2].text = "$150.00"
row[3].text = "0%"
row[4].text = "$1,500.00"

doc.save('/invoices/invoice_2024.docx')
```

---

## System Dependencies

### Required Tools (Optional Features)

| Feature | Dependency | Status |
|---------|------------|--------|
| Create documents | python-docx | ✅ Installed |
| Convert .doc → .docx | LibreOffice | ✅ Available at `/usr/bin/soffice` |
| Extract TOC with redlines | pandoc | ✅ Available at `/usr/bin/pandoc` |
| PDF to images | poppler-utils | ✅ Available at `/usr/bin/pdftoppm` |

### Installation Commands (if needed)

```bash
# All dependencies are already installed on your system!
sudo apt-get install libreoffice poppler-utils  # if not root

# python-docx is in the skills venv and ready to use
```

---

## Comparison with Google Docs Skill

### Word Documents (.docx)

Use `microsoft-word` for:
- Native Microsoft Word files
- Complex tables with specific styling
- Office-specific features like tracked changes
- Legacy .doc compatibility (requires LibreOffice)

### Google Docs / Sheets

Use `google-workspace` for:
- Online collaborative editing
- Google Drive integration
- Cloud-based document management

---

## Error Handling

**Common issues and solutions:**

```python
from hermes_skills import microsoft_word as mw

# Handle missing files
try:
    doc = mw.read_document('/path/to/missing.docx')
except FileNotFoundError:
    print("Document not found. Would you like to create it?")

# Handle invalid formats
if not path.endswith('.docx'):
    print("Please use .docx format for Microsoft Word operations.")
    
# Convert .doc files first
if file_path.endswith('.doc'):
    subprocess.run([
        'soffice', '--headless', 
        '--convert-to', 'docx', file_path, '-outdir'
        '.converted/'
    ])
    path = f"{file_path.rsplit('.', 1)[0]}.docx"
```

---

## Tips & Best Practices

### 1. Use Absolute Paths

Always use absolute paths to avoid confusion between WSL and Windows:

```python
# ✅ Good
path = '/mnt/c/Users/YourName/Documents/report.docx'

# ❌ Avoid
path = './report.docx'  # Relative paths can be ambiguous
```

### 2. Preserve Formatting When Editing

When modifying existing documents, preserve original styles:

```python
doc = microsoft_word.read_document('/original.docx')
# Work with doc.paragraphs[i] and doc.table[j] directly
# Don't add/remove entire paragraphs unless necessary
```

### 3. Image Optimization

Large images can slow down document processing:

```python
# Compress images before adding to documents
optimized_img = compress_image('/large_chart.png', target_size=500*500)
doc.add_image(path='/small_version.png')
```

### 4. Table Width Consistency

Always specify table widths explicitly for consistent rendering:

```python
table = doc.add_table(rows=3, cols=2)
table.style = "Table Grid"  # Default styling
# OR specify width
table.width = {size: 8640, type: "DXA"}  # ~1 inch margins on Letter paper
```

---

## API Reference

### microsoft_word.new_document(**kwargs)

Create a new document from scratch.

**Parameters:**
- `path` (str): File path for the output .docx file
- `title` (str, optional): Document title
- `author` (str, optional): Document author name
- `created` (datetime, optional): Creation timestamp
- `page_size` (str): "Letter" (default) or "A4"
- `orientation` (str): "portrait" (default) or "landscape"

**Returns:** A Document object with methods for adding content.

---

### microsoft_word.read_document(path)

Read an existing document and return its content.

**Parameters:**
- `path` (str): Path to the .docx file to read

**Returns:** A Document object containing:
- `paragraphs`: List of all paragraphs with text and styles
- `tables`: List of all tables with cells and data
- `metadata`: Dictionary with author, created date, etc.
- `text_content`: Plain text of entire document

---

### microsoft_word.Document Object Methods

- `.add_heading(text, level=1)` - Add a styled heading
- `.add_paragraph(text)` - Add plain text paragraph
- `.add_image(path|base64, caption="None")` - Insert image
- `.add_table(rows, cols, headers=None)` - Create table
- `.save()` - Save to file
- `.styles` - Dictionary of available styles for customization

---

### microsoft_word.Table Object Methods

- `table[i][j].text = "..."` - Set cell text
- `table.style = "..."` - Change table style
- `table.width = size` - Set table width

---

## Migration Guide from Python-docx

If you were using plain python-docx:

```python
# OLD (raw python-docx)
from docx import Document as DWX

doc = DWX()
doc.add_heading("Hello", level=0)
doc.save('/tmp/test.docx')

# NEW (microsoft-word skill wrapper)
from hermes_skills import microsoft_word as mw

doc = mw.new_document(
    path='/tmp/test.docx',
    title="Hello"
)
doc.save()  # Automatic saving with metadata
```

The `microsoft-word` skill provides a more intuitive interface while maintaining full compatibility with underlying docx operations.

---

## Examples by Use Case

### Resume Builder

```python
resume = microsoft_word.new_document(
    path='/resumes/my_resume.docx',
    title="Resume Template"
)

resume.add_heading("John Doe", level=0)
resume.add_paragraph("Software Engineer | 10 Years Experience")

resume.add_heading("Professional Summary", level=1)
resume.add_paragraph("""Experienced software engineer specializing in...""")

resume.add_heading("Experience", level=1)
section = resume.add_section()
section.text = """Senior Developer at Company ABC (2020-2024)\n"""
section.text += "Led team of 5 developers...\n"
# ... more experience

resume.save()
```

### Meeting Minutes Template

```python
minutes = microsoft_word.new_document(
    path='/templates/meeting_minutes.docx'
)

title = minutes.add_heading("MEETING MINUTES", level=0)
title.runs[0].font.size = Pt(24)

# Add structured sections with headers for each meeting
def add_meeting_section(doc, title):
    doc.add_heading(title, level=1)
    doc.add_paragraph("Attendees: _________________")
    doc.add_paragraph("Date: _________________")
    
add_meeting_section(minutes, "Project X Planning")
add_meeting_section(minutes, "Q3 Strategy Review")

minutes.save()
```

### Form with Fillable Fields

```python
form = microsoft_word.new_document(
    path='/forms/contact_form.docx'
)

form.add_heading("Contact Form", level=0)

fields = [
    ("Name", "Text"),
    ("Email", "Email"),
    ("Phone", "Phone"),
    ("Company", "Text"),
    ("Message", "Rich Text")
]

for field in fields:
    form.add_heading(field[0], level=2)
    form.add_paragraph(f"Please enter {field[0]}...")
    
form.save()
```

---

## Integration with Other Skills

### Combine with web_search for Reports

```python
from hermes_skills import microsoft_word, web_search

topic = "Sustainability in Tech Industry 2024"
research = web_search.search(query=topic)

report = microsoft_word.new_document(
    path='/reports/sustainability_report.docx',
    title=f"{topic} Report"
)

for snippet in research.results[:5]:  # Top 5 results
    report.add_heading(snippet.title, level=1)
    report.add_paragraph(snippet.summary)

report.save()
```

### Combine with OCR for Scanned Documents

```python
from hermes_skills import ocr_and_documents, microsoft_word

scanned = ocr_and_documents.scan('/path/to/scanned_image.png')
extracted_text = ocr_and_documents.extract_text(scanned_image)

formatted_doc = microsoft_word.new_document(path='/output/ocr_output.docx')
formatted_doc.add_heading("Scanned Document", level=0)
formatted_doc.add_paragraph(extracted_text)

formatted_doc.save()
```

---

## Testing Your Installation

Run these quick tests to verify the skill works:

```python
from hermes_skills import microsoft_word

# Test 1: Create a document
doc = microsoft_word.new_document(path='/tmp/test.docx')
doc.add_heading("Test Document", level=0)
doc.save()
print("✓ Test 1 passed")

# Test 2: Read it back
doc_read = microsoft_word.read_document('/tmp/test.docx')
print(f"Read back: {doc_read.title}")
print("✓ Test 2 passed")

# Test 3: Create a table
doc2 = microsoft_word.new_document(path='/tmp/table_test.docx')
table = doc2.add_table(rows=2, cols=2, headers=["A", "B"])
for i in range(2):
    for j in range(2):
        table[i][j].text = f"{i},{j}"
doc2.save()
print("✓ Test 3 passed")

print("\nAll tests passed! The microsoft-word skill is working.")
```

---

## ⚠️ Deprecation Notice

**Use the `docx` skill directly instead.** The `microsoft-word` skill was created as an alternative interface but has been **superseded by the official `docx` skill**.

### Why Use `docx` Instead?

- ✅ **Same capabilities**: Both skills handle .docx creation/editing
- ✅ **Better integration**: `docx` has comprehensive documentation for both Python and docx-js approaches
- ✅ **System tool awareness**: `docx` documents optional dependencies (pandoc, LibreOffice) with their purposes
- ✅ **Standardization**: Single skill reduces confusion about which API to use

### Migration Guide

**Before** (microsoft-word):
```python
from hermes_skills import microsoft_word
doc = microsoft_word.new_document(...)
```

**After** (docx):
```python
from hermes_skills import docx
doc = docx.new_document(...)  # or read for existing docs
```

The underlying functionality is identical - both use the same Python-docx library.

### When to Use What

| Task | Recommended Skill | Reason |
|------|------------------|--------|
| Create/edit .docx files | `docx` | Full documentation, examples, best practices |
| Table of contents / redlines | `docx` | Requires pandoc (documented) |
| Convert .doc→.docx | `docx` | LibreOffice conversion documented |
| Quick simple docs | Either | Same underlying functionality |

**Bottom line**: Use `docx` - it's the official Anthropic skill with complete documentation and examples.

---

## Summary (Historical)

The `microsoft-word` skill was created as an alternative interface for Word documents but has been deprecated in favor of the comprehensive `docx` skill which provides:
- Full Python docx-js examples for professional documents
- XML unpacking/packing workflows for tracked changes
- System dependency documentation (pandoc, LibreOffice)
- Best practices and design patterns

Simply mention ".docx" or "Word document" to use the `docx` skill directly.