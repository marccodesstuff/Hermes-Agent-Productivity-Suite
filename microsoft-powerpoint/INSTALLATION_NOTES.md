# PPTX Skill Installation Notes

## ✅ What's Installed

- **Skill Location**: `/home/marc/.hermes/skills/productivity/pptx/`
- **Python Environment**: `/home/marc/.hermes/skills/pptx-env/`
- **npm Package**: `pptxgenjs` (installed globally)

## 📦 Functionality by Feature

### ✅ Fully Functional

| Feature | Status | Notes |
|---------|--------|-------|
| Create presentations from scratch | ✓ Works | Using pptxgenjs npm package |
| Read .pptx content | ⚠️ Limited | Requires markitdown[pptx] for text extraction |
| Thumbnail generation | ✓ Works | Pillow installed for thumbnail grids |
| Edit existing presentations | ⚠️ Partial | XML manipulation scripts available |

### ✅ System Dependencies (Available)

| Tool | Purpose | Status |
|------|---------|--------|
| pandoc | Text conversion | ✓ Available at `/usr/bin/pandoc` |
| soffice (LibreOffice) | PPTX processing | ✓ Available at `/usr/bin/soffice` |
| pdftoppm | PDF images | ✓ Available at `/usr/bin/pdftoppm` |

## 🔧 How to Use in Hermes

1. **Direct usage**: Mention "PowerPoint", ".pptx", or "presentation"
2. **API pattern**: `skill_view(name='pptx', path='/path/to/file.pptx')`

The agent will automatically select the appropriate skill based on context.

## 📜 Examples

### Create Presentation from Scratch

```python
from hermes_skills import pptx

pres = pptx.new_presentation(
    path='/home/user/pitchdeck.pptx',
    title="Company Pitch Deck"
)

# Add slides
pres.add_slide(text="Our Vision", layout="title")
pres.add_slide(text="Market Analysis", layout="content")
pres.add_slide(text="Q4 Goals", layout="stats")

pres.save()
```

### Read Existing Presentation

```python
from hermes_skills import pptx

pres = pptx.read_presentation('/path/to/presentation.pptx')
print(f"Slides: {len(pres.slides)}")
print(f"Title: {pres.title}")
```

## 📋 Dependencies Status

- **markitdown[pptx]**: ✓ Installed (text extraction)
- **Pillow**: ✓ Installed (thumbnail generation)
- **pptxgenjs**: ✓ Installed globally via npm (create from scratch)

## ⚠️ Known Limitations

1. **Reading presentations**: Full text extraction requires markitdown[pptx] with pptx support
2. **Editing slides**: XML structure manipulation is more complex than Word documents
3. **Animations/Transitions**: These are stored in different XML files and may need special handling

## 📝 Testing

Test your installation:

```python
from hermes_skills import pptx

# Create a simple presentation
pres = pptx.new_presentation(path='/tmp/test.pptx')
pres.add_slide(text="Hello World", layout="title")
pres.save()

print("✓ PPTX skill is working!")
```

## 📚 Documentation

- **SKILL.md**: Full documentation at `/home/marc/.hermes/skills/productivity/pptx/SKILL.md`
- **Design guides**: See `design-ideas.md` for presentation best practices
- **Editing workflow**: Check `editing.md` for slide manipulation instructions
- **pptxgenjs guide**: Refer to `pptxgenjs.md` for creating from scratch

## 🔒 License

Proprietary - See LICENSE.txt in skill directory for full terms.
