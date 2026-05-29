---
name: adobe-ppt-generator
description: >
  Use this skill whenever creating, editing, or reviewing a PowerPoint
  presentation that should follow Adobe Creative Cloud brand guidelines.
  Triggers include: any mention of "Adobe brand", "Adobe presentation",
  "brand guidelines", "CC brand compliance", requests to build slides
  "for Adobe", or when the user wants a branded deck, pitch, or slide set.
  Also trigger when reviewing an existing .pptx for Adobe brand compliance,
  or when the user says "make me slides for my Adobe plugin/project/launch".
version: 1.0.0
allowed-tools: [Read, Write, Bash]
---

# Adobe-Branded PowerPoint Generator

You create fully brand-compliant Adobe PowerPoint presentations using `python-pptx`.
Apply every rule below — do not skip any section. When in doubt, ask before generating.

---

## Phase 1 — Discovery (ask before writing any code)

Gather these in one message:

1. **Purpose** — What is this deck for? (pitch, internal update, product launch, event, tutorial, sales)
2. **Audience** — Internal Adobians, external partners, customers, developers?
3. **Slides needed** — How many slides? Do they have an outline, or should you suggest one?
4. **Content** — What text, data, or talking points should appear?
5. **Product context** — Which Adobe product/cloud is featured? (determines color world — see §Color below)
6. **Output path** — Where should the .pptx file be saved?

If the user's opening message already answers these, acknowledge and skip straight to generation.

---

## Phase 2 — Slide outline

Propose a slide outline before generating. Typical structure:

| # | Slide type | Purpose |
|---|---|---|
| 1 | Title slide | Hook + presenter name |
| 2 | Agenda | Set expectations |
| 3–N | Content slides | One idea per slide |
| N+1 | Call to action | Next steps |
| N+2 | Thank you / Q&A | Close |

Adjust to the user's request. Get approval before proceeding.

---

## Phase 3 — Generation

Use `python-pptx`. Install if missing: `pip install python-pptx`.

### Slide dimensions

Always use **widescreen 16:9**: width = 33528000 EMU (13.33 in), height = 19051200 EMU (7.5 in).

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Emu(33528000)
prs.slide_height = Emu(19051200)
```

---

## Brand rules — apply to EVERY slide

### Colors

#### Core palette (primary use only)

| Name | HEX | RGBColor |
|---|---|---|
| Adobe Red | `#EB1000` | `RGBColor(0xEB, 0x10, 0x00)` |
| Black | `#000000` | `RGBColor(0x00, 0x00, 0x00)` |
| White | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` |

#### Accent colors (graphics, diagrams, icons, backgrounds ONLY — never text)

| Name | HEX | RGBColor |
|---|---|---|
| Pink | `#FF4885` | `RGBColor(0xFF, 0x48, 0x85)` |
| Magenta | `#F24CB8` | `RGBColor(0xF2, 0x4C, 0xB8)` |
| Purple | `#C844DC` | `RGBColor(0xC8, 0x44, 0xDC)` |
| Violet | `#9A47E2` | `RGBColor(0x9A, 0x47, 0xE2)` |
| Indigo | `#7155FA` | `RGBColor(0x71, 0x55, 0xFA)` |
| Blue | `#3B62FB` | `RGBColor(0x3B, 0x62, 0xFB)` |
| Sky blue | `#1D95E7` | `RGBColor(0x1D, 0x95, 0xE7)` |
| Teal | `#0FB1C0` | `RGBColor(0x0F, 0xB1, 0xC0)` |
| Sea green | `#0DB595` | `RGBColor(0x0D, 0xB5, 0x95)` |
| Green | `#0BA45D` | `RGBColor(0x0B, 0xA4, 0x5D)` |
| Lime | `#5DB41F` | `RGBColor(0x5D, 0xB4, 0x1F)` |
| Yellow-green | `#A3C400` | `RGBColor(0xA3, 0xC4, 0x00)` |
| Gold | `#F5C700` | `RGBColor(0xF5, 0xC7, 0x00)` |
| Orange | `#FFA213` | `RGBColor(0xFF, 0xA2, 0x13)` |

#### Color rules (strict)

- Do NOT use tones, transparencies, or tints of any accent color.
- Do NOT let one accent color dominate the slide — accent = secondary support only.
- **Do NOT use red for data/chart elements** (negative financial connotation).
- Chart semantics: blue = informative, green = positive, orange = notice/caution.

#### Product color worlds

| Product | Primary color world |
|---|---|
| Adobe (corporate) | Red |
| Experience Cloud | Red (Pop or Focus only) |
| Acrobat | Red (Pop or Focus only) |
| Photoshop | Blue |
| Adobe Express | Full-spectrum gradient |

---

### Typography

**Adobe Clean** is the only permitted typeface. Use it for all text boxes.

In `python-pptx`, set `tf.font.name = "Adobe Clean"` on every run.
Fallback chain for rendering environments without the font: `"Segoe UI"`, `"Helvetica Neue"`, `sans-serif`.

#### Weight → PowerPoint font weight mapping

| Role | Weight name | bold= | Pt size (presentation) |
|---|---|---|---|
| Title slide title | Black (900) | True | 60–80 pt |
| Title slide subtitle | Regular (400) | False | 24–36 pt |
| Content slide page title | Black (900) | True | 40–54 pt |
| Section headline | ExtraBold (800) | True | 28–36 pt |
| Subheadline | Bold (700) | True | 20–28 pt |
| Body copy | Regular (400) | False | 16–20 pt |
| Caption / label | Regular (400) | False | 12–14 pt |

*(python-pptx has no native 800/900 weight — use `bold=True`; the Adobe Clean font file itself renders at the correct weight.)*

#### Typography rules

- Sentence-style capitalization on all headings and subheadings.
- Headlines and subheadings end with a **period** or **question mark**.
- Do NOT use ALL CAPS (unless it is an acronym).
- Tracking: set `tf.font.language_id` is fine; `python-pptx` does not expose tracking — note this limitation in a comment.

---

### Slide backgrounds

| Slide type | Background | Text color |
|---|---|---|
| Title slide | Adobe Red | White |
| Section divider | Black | White |
| Content slide | White | Black |
| Dark accent slide | Black | White |

Apply background with `slide.background.fill.solid(); slide.background.fill.fore_color.rgb = RGBColor(...)`.

---

### Text color on backgrounds

| Background | Headlines | Body |
|---|---|---|
| Adobe Red | White | White |
| White | Black | Black |
| Black | White | White |

---

### Logo placement

- Place the **Adobe Wordmark** (white version on Red/Black, Red version on White) in the **bottom-right corner** of every slide.
- Minimum size: 16 px height on screen. In EMU: ~152400 EMU (0.16 in height minimum).
- Clear space = half the width of the Adobe icon on all sides.
- Logo is an image added via `slide.shapes.add_picture(logo_path, left, top, width, height)`.
- If no logo file is available, add a placeholder text box with "ADOBE" in the correct color and note to replace with official asset.
- **Never** recreate the wordmark by setting type yourself — use the approved asset file.

---

### Layout per slide type

#### Title slide

```
Background: Adobe Red
┌─────────────────────────────────────┐
│                                     │
│   [Title — Adobe Clean Black 72pt   │
│    White, sentence case, ends .]    │
│                                     │
│   [Subtitle — Regular 28pt White]   │
│                                     │
│   [Presenter | Date | Event]        │
│                               [LOGO]│
└─────────────────────────────────────┘
```

#### Content slide

```
Background: White
┌─────────────────────────────────────┐
│ [Page title — Black 44pt, Red left  │
│  accent bar 4pt wide]               │
│─────────────────────────────────────│
│                                     │
│  [Body content — Regular 18pt Black]│
│                                     │
│                               [LOGO]│
└─────────────────────────────────────┘
```

Add a 4pt left accent bar in Adobe Red (`add_shape(MSO_SHAPE_TYPE.RECTANGLE, ...)`) for content slides.

#### Section divider

```
Background: Black
┌─────────────────────────────────────┐
│                                     │
│   [Section title — Black 60pt White]│
│                               [LOGO]│
└─────────────────────────────────────┘
```

---

## Python code template

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import copy

ADOBE_RED   = RGBColor(0xEB, 0x10, 0x00)
ADOBE_BLACK = RGBColor(0x00, 0x00, 0x00)
ADOBE_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def set_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, text, left, top, width, height,
                 font_size, bold, color, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Adobe Clean"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_logo(slide, logo_path, slide_width, slide_height):
    logo_h = Emu(457200)   # ~0.48 in
    logo_w = Emu(1143000)  # ~1.19 in (wordmark aspect ratio ~2.5:1)
    margin = Emu(228600)   # ~0.24 in
    slide.shapes.add_picture(
        logo_path,
        slide_width - logo_w - margin,
        slide_height - logo_h - margin,
        logo_w,
        logo_h,
    )


def make_title_slide(prs, title, subtitle, logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    W, H = prs.slide_width, prs.slide_height
    set_bg(slide, ADOBE_RED)
    add_text_box(slide, title,
                 Emu(914400), Emu(2743200), W - Emu(1828800), Emu(3200000),
                 72, True, ADOBE_WHITE, PP_ALIGN.LEFT)
    add_text_box(slide, subtitle,
                 Emu(914400), Emu(5943600), W - Emu(1828800), Emu(914400),
                 28, False, ADOBE_WHITE, PP_ALIGN.LEFT)
    if logo_path:
        add_logo(slide, logo_path, W, H)


def make_content_slide(prs, title, body_lines, logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    W, H = prs.slide_width, prs.slide_height
    set_bg(slide, ADOBE_WHITE)

    # Red accent bar (left edge)
    bar = slide.shapes.add_shape(1, Emu(457200), Emu(685800), Emu(57150), Emu(914400))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ADOBE_RED
    bar.line.fill.background()

    add_text_box(slide, title,
                 Emu(600000), Emu(685800), W - Emu(1200000), Emu(914400),
                 40, True, ADOBE_BLACK, PP_ALIGN.LEFT)

    body = "\n".join(body_lines)
    add_text_box(slide, body,
                 Emu(600000), Emu(1828800), W - Emu(1200000), H - Emu(2743200),
                 18, False, ADOBE_BLACK, PP_ALIGN.LEFT)

    if logo_path:
        add_logo(slide, logo_path, W, H)


def make_section_slide(prs, section_title, logo_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height
    set_bg(slide, ADOBE_BLACK)
    add_text_box(slide, section_title,
                 Emu(914400), Emu(3200000), W - Emu(1828800), Emu(2743200),
                 60, True, ADOBE_WHITE, PP_ALIGN.LEFT)
    if logo_path:
        add_logo(slide, logo_path, W, H)


# --- Build the presentation ---
prs = Presentation()
prs.slide_width  = Emu(33528000)
prs.slide_height = Emu(19051200)

LOGO = None  # Set to file path of approved Adobe Wordmark asset

make_title_slide(prs, "Your title.", "Subtitle or presenter name", LOGO)
make_section_slide(prs, "Section one.", LOGO)
make_content_slide(prs, "Slide title.", ["Bullet one", "Bullet two", "Bullet three"], LOGO)

prs.save("adobe-presentation.pptx")
print("Saved: adobe-presentation.pptx")
```

---

## Phase 4 — Brand compliance check

After generating, run through this checklist and report results to the user:

- [ ] Core palette only (Red `#EB1000`, Black, White) for all primary elements
- [ ] Accent colors used only for graphics/icons/diagrams — never body text or headlines
- [ ] No tones, tints, or transparencies of accent colors
- [ ] Adobe Clean font applied to all text runs
- [ ] Weight hierarchy: Black (title) > ExtraBold/Bold (section/sub) > Regular (body)
- [ ] Sentence-style capitalization; headlines end with period or question mark
- [ ] Text color matches background (white on Red/Black; black on White)
- [ ] Logo placed bottom-right, minimum 16 px height, correct color variant
- [ ] Logo has required clear space (≥ half icon width on all sides)
- [ ] No red used for chart/data elements
- [ ] First mention of Adobe products includes "Adobe" prefix (e.g. "Adobe Photoshop")
- [ ] No TM or ® symbols on Adobe marks
- [ ] Brand voice: human, inspiring, progressive, creative — no jargon

Report any violations and offer to fix them before delivering the file.

---

## Phase 5 — Delivery

1. Save the file to the path the user specified.
2. Confirm the file was created: `ls -lh <path>`.
3. Print the brand compliance checklist results.
4. Offer to open the file or adjust any slide.
