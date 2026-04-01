# Snap to Slide

Turn a photo or video frame of handwritten notes, diagrams, or mixed content into a PowerPoint slide — instantly, using Claude Vision.

## Setup

### 1. Install dependencies

```bash
python3 -m venv venv
source venv/bin/activate  # Windows: .\venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Run the server

```bash
uvicorn server:app --reload --port 8000
```

Then open [http://localhost:8000](http://localhost:8000) in your browser.

## How it works

1. Upload a photo (JPG, PNG, WEBP, GIF) **or a video** (MP4, MOV, WEBM)
2. For videos: drag the frame slider to pick the right moment, then click **Capture this frame**
3. Optionally add a **hint** to guide Claude (e.g. "this is a process flow" or "focus on the numbers")
4. Enter your Anthropic API key and click **Generate**
5. Watch live progress and a preview of the extracted content
6. Download your `.pptx` file

## How slides are built

Claude Vision analyzes the image and decomposes it into **individual visual layers** — text blocks, geometric shapes, arrows, and stat cards — each with a precise normalized position. These layers are rendered back-to-front as native PowerPoint objects (not screenshots), so the result is fully editable in PowerPoint.

Every generated file is validated before download: the PPTX is reopened and checked as a ZIP archive and as a valid OOXML document. If anything is malformed, a readable fallback slide is returned instead of a corrupt file.

## Supported element types

Claude recognises these element types and places them as native PPTX objects:

| Type | Description |
|---|---|
| `title` | Primary slide heading |
| `subtitle` | Secondary heading, italic by default |
| `text_block` | Paragraphs or bullet lists |
| `stat` | Large KPI / metric value with a caption |
| `icon` | Symbol or icon described in text |
| `arrow` | Directed line connector with optional label |
| `shape_rectangle` | Filled rectangle |
| `shape_rounded_rect` | Rounded rectangle |
| `shape_oval` | Ellipse / circle |
| `shape_diamond` | Diamond / rhombus |
| `shape_triangle` | Isosceles triangle |
| `shape_parallelogram` | Parallelogram |
| `shape_trapezoid` | Trapezoid |
| `shape_rectangle_outline` | Rectangle with border, no fill |
| `shape_rounded_rect_outline` | Rounded rectangle with border, no fill |
| `shape_oval_outline` | Ellipse with border, no fill |
| `shape_chevron` | Chevron — mid-sequence process step |
| `shape_home_plate` | Pentagon — first step in a process flow |
| `shape_arrow_right/left/up/down` | Solid block arrow shapes |

To **add a new shape type**, add one entry to `SHAPE_CATALOG` in `server.py` — no other code changes needed.

## Slide themes

The UI offers a dropdown of named themes. Each theme controls the slide background, accent colour, and font pairing (heading + body).

| Theme | Style |
|---|---|
| Default | Warm off-white, terracotta accent, Georgia / Trebuchet MS |
| Atlas | Dark navy, sky-blue accent, Calibri |
| Celestial | Deep space blue, purple accent, Palatino |
| Madison | Clean white, deep navy accent, Garamond |
| Retrospect | Light grey, red accent, Rockwell |
| Slate | Dark charcoal, teal accent, Trebuchet MS |
| Corporate | White, corporate blue, Arial |
| Organic | Sage green, forest accent, Georgia / Verdana |
| Solarized | Warm parchment, gold accent, Consolas |

### Using real PowerPoint template files

For full fidelity — gradients, background art, master layouts — drop a `.pptx` file into the `templates/` directory, named exactly after the theme (e.g. `Atlas.pptx`). The server will use it as the base presentation and add content on top.

**How to export a theme from PowerPoint:**
1. Open PowerPoint → *Design* → choose a theme
2. *File → Save As → PowerPoint Presentation* (`.pptx`)
3. Copy the saved file to `templates/<ThemeName>.pptx`

**Free template sources (no Office licence required):**
- [SlidesCarnival](https://www.slidescarnival.com) — free, CC-licensed `.pptx` files
- [SlidesMania](https://slidesmania.com) — free Google Slides / PPTX templates
- [FPPT](https://www.free-powerpoint-templates-design.com) — large free library

Themes that have a matching file show a **✦** marker in the dropdown.

To add a **new theme**, add one entry to `SLIDE_THEMES` in `server.py` and optionally drop a matching `.pptx` into `templates/`.

## Project structure

```
snap-to-slide/
├── server.py          # FastAPI backend
├── requirements.txt
├── templates/         # Optional .pptx theme base files (user-provided)
├── static/
│   ├── index.html     # HTML shell
│   ├── style.css      # All styles
│   └── app.js         # All client-side logic
└── README.md
```
