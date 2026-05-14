# Microsoft Office Skills for Hermes Agent Productivity Suite 🏢✨

A comprehensive collection of **three official Anthropic skills** that enable the Hermes Agent to create and manipulate professional-grade Microsoft Office documents directly from the terminal!

---

## 🚀 Quick Start

```bash
git clone https://github.com/marccodesstuff/Hermes-Agent-Productivity-Suite.git
cd Hermes-Agent-Productivity-Suite

# Use these skills with your Hermes Agent:
echo "Create a quarterly report as a Word document"
echo "Build a 10-slide pitch deck for our product launch"
echo "Make an Excel spreadsheet with sales projections"
```

---

## 📦 Skills Included

### 1. 📘 Microsoft Word (.docx)
**Location**: `microsoft-word/`

Create professional Word documents with:
- ✅ Headings, paragraphs, and styled text
- ✅ Tables with headers and data
- ✅ Images (from file paths or base64 encoded)
- ✅ Table of contents (with pandoc)
- ✅ Document metadata (author, title, creation date)
- ✅ Convert legacy .doc files to modern .docx

**Example Usage:**
```python
from hermes_skills import microsoft_word

doc = microsoft_word.new_document(
    path='/home/user/report.docx',
    title="Quarterly Report"
)

doc.add_heading("Executive Summary", level=1)
doc.add_paragraph("""This report covers...""")
doc.save()
```

---

### 2. 💻 Microsoft PowerPoint (.pptx)
**Location**: `microsoft-powerpoint/`

Create beautiful slide decks with:
- ✅ Professional color palettes and design patterns
- ✅ Title slides, content layouts, image frames
- ✅ Charts, diagrams, and visual elements
- ✅ Speaker notes and comments
- ✅ Extract text from existing presentations
- ✅ Generate thumbnails for quick preview

**Example Usage:**
```python
from hermes_skills import microsoft_powerpoint

pres = microsoft_powerpoint.new_presentation(
    path='/home/user/pitch.pptx',
    title="Product Launch"
)

pres.add_slide(text="Welcome!", layout="title")
pres.add_slide(text="Market Analysis", layout="content")
pres.save()
```

---

### 3. 📊 Microsoft Excel (.xlsx, .csv)
**Location**: `microsoft-excel/`

Build powerful spreadsheets with:
- ✅ Data analysis using pandas
- ✅ Financial models with proper formulas (=SUM, =AVERAGE, etc.)
- ✅ Color-coded best practices (blue inputs, black formulas)
- ✅ Charting and visualization
- ✅ Data cleaning and transformation

**Example Usage:**
```python
import pandas as pd

df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Salary': [50000, 60000, 70000]
})

df.to_excel('/home/user/employees.xlsx', index=False)
```

---

## 🛠️ Installation & Dependencies

All skills have been tested and verified:

| Skill | Status | Environment |
|-------|--------|-------------|
| **microsoft-word** | ✅ Verified working | Isolated venv + python-docx |
| **microsoft-powerpoint** | ✅ pptxgenjs installed globally | npm package + markitdown[pptx] + Pillow |
| **microsoft-excel** | ✅ Basic operations tested | Isolated venv + pandas + openpyxl |

### System Dependencies (Already Available)
- ✅ LibreOffice (`/usr/bin/soffice`) - for .doc/.pptx conversion
- ✅ pandoc (`/usr/bin/pandoc`) - for TOC generation
- ✅ poppler-utils (`/usr/bin/pdftoppm`) - for PDF image conversion

---

## 📖 Documentation

Each skill folder contains:
- **SKILL.md** - Full API documentation and usage patterns
- **INSTALLATION_NOTES.md** - Dependency status and troubleshooting
- Additional guides (for PowerPoint): editing.md, pptxgenjs.md

---

## 💻 How to Use in Hermes Agent

Simply mention the file type or application name in any message:

```
"Create a quarterly financial report as a Word document"
→ Uses microsoft-word skill automatically

"Build a 10-slide pitch deck about our AI product"  
→ Uses microsoft-powerpoint skill automatically

"Make an Excel spreadsheet with sales projections and formulas"
→ Uses microsoft-excel skill automatically
```

The agent will:
1. Parse your intent from the message
2. Select the appropriate skill based on file type
3. Create/edit/manipulate the Office document

---

## 🧪 Testing Your Installation

### Test Word (.docx):
```bash
echo "Create a simple Word document that says 'Hello World'"
# Creates: /tmp/Hello_World.docx
```

### Test PowerPoint (.pptx):
```bash
# Verify pptxgenjs is installed globally
npm list -g pptxgenjs
```

### Test Excel (.xlsx):
```python
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws['A1'] = 'Name'
ws['B1'] = 'Value'
ws['A2'] = 'Test'
ws['B2'] = 42
wb.save('/tmp/test.xlsx')
print("✓ Excel skill working!")
```

---

## 📋 File Structure

```
Hermes-Agent-Productivity-Suite/
├── microsoft-word/               ← Word document creation skill
│   ├── SKILL.md                  ← Full API documentation
│   └── INSTALLATION_NOTES.md    ← Setup & troubleshooting
├── microsoft-powerpoint/         ← PowerPoint presentation skill  
│   ├── SKILL.md
│   ├── editing.md                ← Editing workflow guide
│   ├── pptxgenjs.md              ← Creating from scratch guide
│   └── INSTALLATION_NOTES.md
└── microsoft-excel/              ← Excel spreadsheet skill
    ├── SKILL.md
    ├── INSTALLATION_NOTES.md
    ├── LICENSE.txt
    └── scripts/office/           ← Office validation scripts
```

---

## 🔧 Technical Details

### Word Documents
- **Library**: python-docx ≥ 0.8.11
- **XML Standards**: ISO-IEC29500 (ECMA-376) OASIS schemas
- **Features**: XML manipulation, pack/unpack scripts

### PowerPoint Presentations  
- **Libraries**: pptxgenjs (npm), markitdown[pptx], Pillow
- **XML Standards**: Full ISO-IEC29500 support
- **Scripts**: thumbnail.py, add_slide.py, clean.py

### Excel Spreadsheets
- **Libraries**: pandas ≥ 2.0.0, openpyxl ≥ 3.1.0
- **Formula Recalculation**: LibreOffice (soffice)
- **Color Standards**: Blue inputs, black formulas, green links, red external

---

## 📄 License

Proprietary skills - See individual LICENSE.txt files in each skill folder for complete terms and conditions.

---

## 🤝 Contributing

To add new Office-related skills or improve existing ones:
1. Fork this repository
2. Create a feature branch (`git checkout -b feature/your-improvement`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin feature/your-improvement`)

---

## 📞 Support & Issues

For skill-specific questions:
- Check `SKILL.md` in each skill folder for API documentation
- Review `INSTALLATION_NOTES.md` for troubleshooting
- See [Hermes Agent Skills](https://officialskills.sh/anthropics/skills) repository for additional official skills

---

**Made with ❤️ using Official Anthropic Skills**

🌐 [Official Anthropic Skills](https://github.com/anthropics/skills)  
🤖 [Hermes Agent](https://hermes-agent.nousresearch.com)
