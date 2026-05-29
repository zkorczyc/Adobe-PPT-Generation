# Adobe PPT Generation — Claude Code Plugin

A Claude Code marketplace plugin that generates fully Adobe-branded PowerPoint presentations (`.pptx`) following the **Adobe Creative Cloud Brand Guidelines (January 2025)**.

## What's inside

| Skill | Purpose |
|---|---|
| **`adobe-ppt-generator`** | End-to-end wizard: discovers requirements, proposes a slide outline, generates brand-compliant `.pptx` via `python-pptx`, and runs a brand compliance check. |

## How it works

The skill guides you through 5 phases:

1. **Discovery** — asks about purpose, audience, slide count, content, and product context
2. **Slide outline** — proposes a structure and waits for approval
3. **Generation** — writes Python code using `python-pptx` to produce the `.pptx` file
4. **Brand compliance check** — validates colors, typography, logo, voice, and layout against Adobe guidelines
5. **Delivery** — saves the file, reports the compliance results, offers revisions

## Brand rules enforced

- **Colors** — Core palette: Adobe Red `#EB1000`, Black, White. Accent colors for graphics only.
- **Typography** — Adobe Clean typeface exclusively; 5 approved weights (Black to SemiLight); sentence-case headlines ending with a period.
- **Layout** — Widescreen 16:9; title slide on Red, content slides on White, section dividers on Black.
- **Logo** — Adobe Wordmark bottom-right, minimum size, correct color variant per background.
- **Voice** — Human, inspiring, progressive, creative. AI copy focuses on the human, not the tool.

## Requirements

- Python 3.8+
- `python-pptx` (`pip install python-pptx`)
- Approved Adobe Wordmark asset (`.png` or `.svg`) for logo placement

## Usage

Ask Claude to create an Adobe-branded presentation:

```
Make me an Adobe-branded deck for my Firefly plugin launch — 8 slides, audience is external developers.
```

Or invoke directly:

```
/adobe-ppt-generation:adobe-ppt-generator
```

## Installation

Add this plugin to your Claude Code project:

```bash
claude plugin add https://github.com/zkorczyc/Adobe-PPT-Generation
```