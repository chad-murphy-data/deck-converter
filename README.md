# Deck Converter

Drag-and-drop tool that converts existing PowerPoint decks into styled templates.

## Quick Start

### One-time setup
```bash
pip install python-pptx flask
```

### Run it
**Option A — Double-click:**  
Double-click `Start Deck Converter.command` (Mac)

**Option B — Terminal:**
```bash
cd deck-converter
python app.py
```

Your browser opens to `http://localhost:5000`.

## How It Works

1. **Drop** your .pptx file onto the browser page
2. **Review** the auto-detected slide types (title, in_brief, comparison, etc.)
3. **Override** any misdetected slides using the dropdowns
4. **Pick** a template style (Slick Minimal or Colorful)
5. **Build** → download your converted deck

## Templates

**Slick Minimal** — Thick green left accent bar, thin rules under titles, clean and understated.

**Colorful** — Colored header bars, multi-color card system with numbered circles, vibrant and energetic.

## Slide Types

| Type | Auto-detects when... |
|------|---------------------|
| Title | First slide, few words, large font |
| Agenda | Contains "agenda", "outline", "overview" |
| In Brief | 3+ bullet-length text items |
| Section Divider | Very short text, large font |
| Stat Callout | Standalone big number/percentage |
| Quote | Text in quotation marks with attribution |
| Comparison | "vs" keywords or two-column layout |
| Text + Graph | Slide contains a chart |
| Process Flow | Numbered steps or "step 1/2/3" patterns |
| Matrix (2×2) | "quadrant", "matrix", "framework" keywords |
| Methods | Multiple methodology keywords |
| Hypotheses | "hypothesis" keywords or confirmed/rejected status |
| WSN Dense | "What / So What / Now What" structure |
| WSN Reveal | Same as above, builds across 3 slides |
| Findings & Recs | Finding/recommendation pairs or → arrows |
| Open Questions | 3+ sentences ending in "?" |
| Progressive Reveal | Multi-point build with running takeaways |
| Closer | "Thank you", "Q&A", short text |

## File Structure

```
deck-converter/
├── app.py                          # Flask web server
├── detector.py                     # Auto-detection engine
├── mapper.py                       # Content → template data mapper
├── template_slick.py               # Slick Minimal builder (python-pptx)
├── template_colorful.py            # Colorful builder (python-pptx)
├── static/index.html               # Browser UI
├── Start Deck Converter.command    # Mac double-click launcher
├── uploads/                        # Temp uploaded files
└── output/                         # Generated decks
```

## Notes

- **Charts:** Chart data from original decks can't be extracted automatically. The converter places a placeholder chart — update the data manually in PowerPoint.
- **Images:** Images from original decks aren't transferred. The tool focuses on text content.
- **Fonts:** Templates use Calibri as a safe fallback. If you have Fidelity Slab/Sans installed, edit the `TITLE_FONT` and `BODY_FONT` constants in the template files.
