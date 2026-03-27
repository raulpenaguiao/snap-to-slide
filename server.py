import asyncio
import base64
import json
import os
import uuid
from io import BytesIO
from typing import Optional

import anthropic
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job store: job_id -> {"image_data": bytes, "mime_type": str}
jobs: dict[str, dict] = {}
# In-memory download store: token -> bytes
downloads: dict[str, bytes] = {}


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert 6-char hex string (no #) to RGBColor."""
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_top_bar(slide, width, accent_hex: str):
    """Add full-width accent bar at top (~5px tall)."""
    bar = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(0), Inches(0), width, Pt(6),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb(accent_hex)
    bar.line.fill.background()


def add_footer(slide, width, height, text_hex: str):
    """Add 'Snap to Slide' footer bottom-left."""
    tf_box = slide.shapes.add_textbox(
        Inches(0.2), height - Pt(20), Inches(3), Pt(20)
    )
    tf = tf_box.text_frame
    tf.text = "Snap to Slide"
    p = tf.paragraphs[0]
    run = p.runs[0]
    run.font.size = Pt(8)
    run.font.color.rgb = hex_to_rgb(text_hex)
    run.font.bold = False


def rect_border_point(cx, cy, hw, hh, dx, dy):
    """Return the point on a rectangle's border in direction (dx, dy) from center.
    cx, cy = center; hw, hh = half-width, half-height; dx, dy = direction."""
    if dx == 0 and dy == 0:
        return cx, cy
    ts = []
    if dx > 0:
        ts.append(hw / dx)
    elif dx < 0:
        ts.append(-hw / dx)
    if dy > 0:
        ts.append(hh / dy)
    elif dy < 0:
        ts.append(-hh / dy)
    t = min(ts)
    return cx + t * dx, cy + t * dy


def build_bullets_slide(prs, data: dict) -> None:
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Left accent bar
    bar = slide.shapes.add_shape(
        1,
        Inches(0.3), Inches(0.5), Pt(5), Inches(4.5),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    bar.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.65), Inches(0.45), Inches(8.8), Inches(0.7)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    # Bullets
    bullets = content.get("bullets", [])
    bullet_box = slide.shapes.add_textbox(
        Inches(0.65), Inches(1.3), Inches(8.8), Inches(3.4)
    )
    tf = bullet_box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = f"• {bullet}"
        run.font.size = Pt(16)
        run.font.color.rgb = hex_to_rgb(theme["text"])

    # Note bar at bottom
    note = content.get("note", "")
    if note:
        note_box = slide.shapes.add_shape(
            1,
            Inches(0), slide_height - Inches(0.55), slide_width, Inches(0.45),
        )
        note_box.fill.solid()
        note_box.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
        note_box.line.fill.background()
        tf2 = note_box.text_frame
        tf2.margin_left = Inches(0.2)
        p = tf2.paragraphs[0]
        run = p.add_run()
        run.text = note
        run.font.size = Pt(10)
        run.font.color.rgb = hex_to_rgb(theme["bg"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_two_column_slide(prs, data: dict) -> None:
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.35), Inches(9.2), Inches(0.65)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    # Horizontal divider
    divider = slide.shapes.add_shape(
        1,
        Inches(0.4), Inches(1.1), Inches(9.2), Pt(2),
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    divider.line.fill.background()

    # Left column
    left_text = content.get("left_column", "")
    left_box = slide.shapes.add_textbox(
        Inches(0.4), Inches(1.25), Inches(4.4), Inches(3.8)
    )
    tf = left_box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(left_text.split("\n") if left_text else []):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(14)
        run.font.color.rgb = hex_to_rgb(theme["text"])

    # Vertical separator
    vsep = slide.shapes.add_shape(
        1,
        Inches(4.95), Inches(1.15), Pt(2), Inches(3.9),
    )
    vsep.fill.solid()
    vsep.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    vsep.line.fill.background()

    # Right column
    right_text = content.get("right_column", "")
    right_box = slide.shapes.add_textbox(
        Inches(5.15), Inches(1.25), Inches(4.4), Inches(3.8)
    )
    tf = right_box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(right_text.split("\n") if right_text else []):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(14)
        run.font.color.rgb = hex_to_rgb(theme["text"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_key_stats_slide(prs, data: dict) -> None:
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.35), Inches(9.2), Inches(0.65)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    stats = content.get("stats", [])
    if not stats:
        add_footer(slide, slide_width, slide_height, theme["text"])
        return

    n = len(stats)
    card_w = Inches(min(2.5, 9.0 / n))
    card_h = Inches(2.0)
    total_w = card_w * n
    start_x = (slide_width - total_w) / 2
    card_y = (slide_height - card_h) / 2 + Inches(0.2)

    for i, stat in enumerate(stats):
        cx = start_x + i * card_w

        # Card background
        card = slide.shapes.add_shape(
            1,
            cx + Inches(0.05), card_y, card_w - Inches(0.1), card_h,
        )
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
        card.line.fill.background()

        # Stat value
        val_box = slide.shapes.add_textbox(
            cx + Inches(0.05), card_y + Inches(0.15),
            card_w - Inches(0.1), Inches(1.1)
        )
        tf = val_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = stat.get("value", "")
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = hex_to_rgb(theme["bg"])

        # Stat label
        lbl_box = slide.shapes.add_textbox(
            cx + Inches(0.05), card_y + Inches(1.3),
            card_w - Inches(0.1), Inches(0.6)
        )
        tf = lbl_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = stat.get("label", "")
        run.font.size = Pt(12)
        run.font.color.rgb = hex_to_rgb(theme["bg"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_title_content_slide(prs, data: dict) -> None:
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.35), Inches(9.0), Inches(0.7)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(30)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    # Subtitle
    subtitle = content.get("subtitle", "")
    if subtitle:
        sub_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.1), Inches(9.0), Inches(0.45)
        )
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = subtitle
        run.font.size = Pt(16)
        run.font.italic = True
        run.font.color.rgb = hex_to_rgb(theme["accent"])

    # Main text panel
    main_text = content.get("main_text", "")
    panel_y = Inches(1.65) if subtitle else Inches(1.2)
    panel_h = Inches(3.0) if content.get("note") else Inches(3.4)

    panel = slide.shapes.add_shape(
        1,
        Inches(0.5), panel_y, Inches(9.0), panel_h,
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
    panel.line.fill.background()
    tf2 = panel.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.2)
    tf2.margin_top = Inches(0.15)
    tf2.margin_right = Inches(0.2)
    p = tf2.paragraphs[0]
    run = p.add_run()
    run.text = main_text
    run.font.size = Pt(15)
    run.font.color.rgb = hex_to_rgb(theme["bg"])

    # Note bar
    note = content.get("note", "")
    if note:
        note_box = slide.shapes.add_shape(
            1,
            Inches(0), slide_height - Inches(0.55), slide_width, Inches(0.45),
        )
        note_box.fill.solid()
        note_box.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
        note_box.line.fill.background()
        tf3 = note_box.text_frame
        tf3.margin_left = Inches(0.2)
        p = tf3.paragraphs[0]
        run = p.add_run()
        run.text = note
        run.font.size = Pt(10)
        run.font.color.rgb = hex_to_rgb(theme["bg"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_diagram_slide(prs, data: dict) -> None:
    """Render a geometry-aware diagram slide: boxes (nodes) connected by arrows (edges)."""
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Content area bounds (nodes' normalized coords map into this region)
    AREA_X = Inches(0.3)
    AREA_Y = Inches(0.9)
    AREA_W = Inches(9.4)
    AREA_H = Inches(4.2)

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.2), Inches(9.2), Inches(0.6)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    nodes = content.get("nodes", [])
    edges = content.get("edges", [])

    # Compute absolute positions and build center map for edge drawing
    node_map: dict[str, dict] = {}
    for node in nodes:
        ax = AREA_X + node["x"] * AREA_W
        ay = AREA_Y + node["y"] * AREA_H
        aw = node["w"] * AREA_W
        ah = node["h"] * AREA_H
        node_map[node["id"]] = {
            "cx": ax + aw / 2,
            "cy": ay + ah / 2,
            "hw": aw / 2,
            "hh": ah / 2,
            "x": ax, "y": ay, "w": aw, "h": ah,
        }

    # Draw edges first (behind nodes)
    for edge in edges:
        fid = edge.get("from")
        tid = edge.get("to")
        if fid not in node_map or tid not in node_map:
            continue
        fn = node_map[fid]
        tn = node_map[tid]
        dx = tn["cx"] - fn["cx"]
        dy = tn["cy"] - fn["cy"]
        # Exit point from source node border
        x1, y1 = rect_border_point(fn["cx"], fn["cy"], fn["hw"], fn["hh"], dx, dy)
        # Entry point at target node border
        x2, y2 = rect_border_point(tn["cx"], tn["cy"], tn["hw"], tn["hh"], -dx, -dy)

        try:
            connector = slide.shapes.add_connector(
                MSO_CONNECTOR_TYPE.STRAIGHT, x1, y1, x2, y2
            )
            connector.line.color.rgb = hex_to_rgb(theme["accent"])
            connector.line.width = Pt(1.5)
            # Add arrowhead at the end using OxmlElement (correct namespace handling)
            ln = connector.line._ln
            if ln is not None:
                tail_end = OxmlElement("a:tailEnd")
                tail_end.set("type", "arrow")
                tail_end.set("w", "med")
                tail_end.set("len", "med")
                ln.append(tail_end)
        except Exception:
            pass  # Connector drawing is best-effort

        # Edge label (if any)
        label = edge.get("label", "")
        if label:
            lx = (x1 + x2) / 2 - Inches(0.4)
            ly = (y1 + y2) / 2 - Pt(8)
            lbl = slide.shapes.add_textbox(lx, ly, Inches(0.8), Pt(16))
            tf = lbl.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = label
            run.font.size = Pt(9)
            run.font.italic = True
            run.font.color.rgb = hex_to_rgb(theme["accent"])

    # Draw nodes (on top of edges)
    for node in nodes:
        nm = node_map[node["id"]]
        shape = slide.shapes.add_shape(
            1,  # MSO_AUTO_SHAPE_TYPE.RECTANGLE (same as original code uses)
            nm["x"], nm["y"], nm["w"], nm["h"],
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
        shape.line.color.rgb = hex_to_rgb(theme["accent"])
        shape.line.width = Pt(1)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.03)
        tf.margin_bottom = Inches(0.03)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = node.get("text", "")
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = hex_to_rgb(theme["bg"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_table_slide(prs, data: dict) -> None:
    """Render a structured table slide."""
    theme = data["theme"]
    content = data["content"]
    title_text = data.get("title", "")

    headers = content.get("headers", [])
    rows = content.get("rows", [])

    if not headers and not rows:
        build_bullets_slide(prs, data)
        return

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    add_top_bar(slide, slide_width, theme["accent"])

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.2), Inches(9.2), Inches(0.6)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["primary"])

    num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)
    num_rows = len(rows) + (1 if headers else 0)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(0.4), Inches(0.9),
        Inches(9.2), Inches(4.3),
    )
    table = table_shape.table

    row_offset = 0
    if headers:
        for j, header in enumerate(headers[:num_cols]):
            cell = table.cell(0, j)
            cell.text = str(header)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.runs[0] if p.runs else p.add_run()
            run.font.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = hex_to_rgb(theme["bg"])
        row_offset = 1

    for i, row in enumerate(rows):
        for j, cell_text in enumerate(list(row)[:num_cols]):
            cell = table.cell(i + row_offset, j)
            cell.text = str(cell_text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(theme["bg"])
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(11)
            run.font.color.rgb = hex_to_rgb(theme["text"])

    add_footer(slide, slide_width, slide_height, theme["text"])


def build_pptx(data: dict) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    layout = data.get("layout", "bullets")
    builders = {
        "bullets": build_bullets_slide,
        "two_column": build_two_column_slide,
        "key_stats": build_key_stats_slide,
        "title_content": build_title_content_slide,
        "diagram": build_diagram_slide,
        "table": build_table_slide,
    }
    builder = builders.get(layout, build_bullets_slide)
    builder(prs, data)

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── Prompts ────────────────────────────────────────────────────────────────────

STRUCTURE_PROMPT = """Analyze this image and identify its primary visual structure.
Return ONLY valid JSON (no markdown, no explanation):
{
  "structure_type": "diagram|table|chart|text",
  "notes": "one sentence describing the spatial layout and key visual features"
}

Structure types:
- "diagram": flowcharts, process flows, org charts, mind maps, network diagrams — shapes/boxes connected by arrows or lines
- "table": grids, matrices, comparison tables, spreadsheet-like rows and columns of data
- "chart": bar charts, pie charts, line graphs, scatter plots — numerical/statistical visualizations
- "text": written content, bullet lists, slides with paragraphs, notes, outlines, or any primarily textual content"""


DIAGRAM_EXTRACTION_PROMPT = """You are analyzing a diagram. Extract its structure as positioned nodes and directed edges, preserving the spatial layout exactly as it appears.

Return ONLY valid JSON:
{
  "title": "descriptive title of the diagram",
  "layout": "diagram",
  "theme": {
    "bg": "f5f0e8",
    "primary": "2c3e50",
    "accent": "c84b31",
    "text": "333333"
  },
  "content": {
    "nodes": [
      {"id": "n1", "text": "Node label", "x": 0.05, "y": 0.3, "w": 0.2, "h": 0.15}
    ],
    "edges": [
      {"from": "n1", "to": "n2", "label": ""}
    ]
  }
}

SPATIAL RULES — these are critical:
- x, y, w, h are normalized 0.0–1.0 relative to the diagram area
- x, y = top-left corner of the node; w, h = width and height of the node
- Mirror the exact spatial layout: if A appears left of B in the image, A must have a smaller x value
- Nodes at the top of the image get small y values; nodes at the bottom get larger y values
- Do NOT cluster all nodes at the same position — spread them to match the original
- Typical node: w=0.15–0.25, h=0.10–0.18; do not make nodes too small (min w=0.10, h=0.08)
- Nodes must not overlap: ensure rectangles [x, x+w] × [y, y+h] do not intersect
- Include ALL visible nodes and ALL arrows/connections
- edges: directed from source to destination; include label text if present on the arrow

Theme: choose colors that complement the content. All hex values 6 chars, NO # symbol."""


TABLE_EXTRACTION_PROMPT = """You are analyzing a table or grid. Extract it precisely as structured data.

Return ONLY valid JSON:
{
  "title": "table title or topic",
  "layout": "table",
  "theme": {
    "bg": "f8f9fa",
    "primary": "1a3a5c",
    "accent": "2980b9",
    "text": "2c3e50"
  },
  "content": {
    "headers": ["Column 1", "Column 2", "Column 3"],
    "rows": [
      ["row1col1", "row1col2", "row1col3"],
      ["row2col1", "row2col2", "row2col3"]
    ]
  }
}

Rules:
- Extract ALL visible rows and columns faithfully — do not summarize or skip rows
- headers: column header texts (empty array [] if there are no headers)
- rows: array of arrays; each inner array is one data row in order
- Preserve exact cell text content; do not paraphrase
- If a cell spans multiple columns, repeat the value across those columns
Theme: all hex values 6 chars, NO # symbol."""


CHART_EXTRACTION_PROMPT = """You are analyzing a chart or graph. Extract the key data points and metrics.

Return ONLY valid JSON:
{
  "title": "chart title",
  "layout": "key_stats",
  "theme": {
    "bg": "f5f0e8",
    "primary": "2c3e50",
    "accent": "c84b31",
    "text": "333333"
  },
  "content": {
    "stats": [
      {"value": "42%", "label": "Metric Name"}
    ],
    "note": "Chart type and key insight in one sentence"
  }
}

Rules:
- Extract the most important 2–5 data points as stats (the standout numbers, peaks, or totals)
- "note": describe the chart type (bar, pie, line…) and the primary trend or finding
- Keep value strings concise (e.g. "42%", "$1.2M", "3.7x")
Theme: all hex values 6 chars, NO # symbol."""


TEXT_EXTRACTION_PROMPT = """You are a PowerPoint slide designer. Analyze the image and extract its content, then design a single professional slide.

Return ONLY valid JSON (no markdown, no explanation) with this exact structure:
{
  "title": "slide title",
  "layout": "bullets",
  "theme": {
    "bg": "f5f0e8",
    "primary": "2c3e50",
    "accent": "c84b31",
    "text": "333333"
  },
  "content": {
    "main_text": "",
    "subtitle": "",
    "bullets": [],
    "left_column": "",
    "right_column": "",
    "stats": [],
    "note": ""
  }
}

Layout rules:
- "bullets": use for notes, lists, outlines — populate "bullets" array and optionally "note"
- "two_column": use for comparisons or two topics — populate "left_column" and "right_column"
- "key_stats": use for numbers, metrics, data — populate "stats" array as [{"value": "42%", "label": "Growth"}, ...]
- "title_content": use for single topic with explanation — populate "title", "subtitle", "main_text", optionally "note"

Theme: choose colors that suit the content. All hex values are 6 chars, NO # symbol.
Extract all visible text. Be faithful to the source content."""

EXTRACTION_PROMPTS = {
    "diagram": DIAGRAM_EXTRACTION_PROMPT,
    "table": TABLE_EXTRACTION_PROMPT,
    "chart": CHART_EXTRACTION_PROMPT,
    "text": TEXT_EXTRACTION_PROMPT,
}


async def process_job(job_id: str):
    """Generator that streams SSE events for a job."""
    job = jobs.get(job_id)
    if not job:
        yield f"data: {json.dumps({'error': 'Job not found'})}\n\n"
        return

    image_data = job["image_data"]
    mime_type = job["mime_type"]
    api_key = job["api_key"]

    try:
        b64 = base64.standard_b64encode(image_data).decode("utf-8")
        client = anthropic.Anthropic(api_key=api_key)
        loop = asyncio.get_event_loop()

        def make_vision_call(prompt_text: str, max_tok: int = 512):
            return client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=max_tok,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": mime_type,
                                "data": b64,
                            },
                        },
                        {"type": "text", "text": prompt_text},
                    ],
                }],
            )

        # ── Step 1: Detect visual structure ───────────────────────────────────
        yield f"data: {json.dumps({'step': 1, 'message': 'Detecting visual structure…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.05)

        struct_response = await loop.run_in_executor(
            None, lambda: make_vision_call(STRUCTURE_PROMPT, 256)
        )
        struct_text = struct_response.content[0].text.strip()
        if struct_text.startswith("```"):
            lines = struct_text.split("\n")
            struct_text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])
        struct_data = json.loads(struct_text)
        structure_type = struct_data.get("structure_type", "text")
        if structure_type not in EXTRACTION_PROMPTS:
            structure_type = "text"

        yield f"data: {json.dumps({'step': 1, 'message': 'Detecting visual structure…', 'status': 'done'})}\n\n"

        # ── Step 2: Extract content with geometry-aware prompt ─────────────────
        type_label = {
            "diagram": "diagram/flowchart",
            "table": "table/grid",
            "chart": "chart/graph",
            "text": "text content",
        }.get(structure_type, "content")
        yield f"data: {json.dumps({'step': 2, 'message': f'Extracting {type_label}…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.05)

        extraction_prompt = EXTRACTION_PROMPTS[structure_type]
        extract_response = await loop.run_in_executor(
            None, lambda: make_vision_call(extraction_prompt, 2048)
        )
        raw_text = extract_response.content[0].text.strip()

        yield f"data: {json.dumps({'step': 2, 'message': f'Extracting {type_label}…', 'status': 'done'})}\n\n"

        # ── Step 3: Parse and validate ─────────────────────────────────────────
        yield f"data: {json.dumps({'step': 3, 'message': 'Designing slide structure…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.05)

        if raw_text.startswith("```"):
            lines = raw_text.split("\n")
            raw_text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])

        slide_data = json.loads(raw_text)

        yield f"data: {json.dumps({'step': 3, 'message': 'Designing slide structure…', 'status': 'done'})}\n\n"

        # ── Step 4: Build PPTX ─────────────────────────────────────────────────
        yield f"data: {json.dumps({'step': 4, 'message': 'Building PowerPoint file…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.05)

        pptx_bytes = await loop.run_in_executor(None, lambda: build_pptx(slide_data))

        token = str(uuid.uuid4())
        downloads[token] = pptx_bytes
        del jobs[job_id]

        yield f"data: {json.dumps({'step': 4, 'message': 'Building PowerPoint file…', 'status': 'done', 'download_token': token})}\n\n"

    except json.JSONDecodeError as e:
        yield f"data: {json.dumps({'error': f'Failed to parse Claude response: {str(e)}'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'error': str(e)})}\n\n"


@app.post("/generate")
async def generate(file: UploadFile = File(...), api_key: str = File(...)):
    """Accept image upload + API key, store them, return job_id."""
    if not api_key or not api_key.strip():
        raise HTTPException(status_code=400, detail="API key is required")
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image")

    image_data = await file.read()
    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "image_data": image_data,
        "mime_type": file.content_type,
        "api_key": api_key.strip(),
    }
    return {"job_id": job_id}


@app.get("/stream/{job_id}")
async def stream(job_id: str):
    """SSE stream for a job."""
    return StreamingResponse(
        process_job(job_id),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/download/{token}")
async def download(token: str):
    """Download the generated .pptx file."""
    pptx_bytes = downloads.get(token)
    if not pptx_bytes:
        raise HTTPException(status_code=404, detail="File not found or already downloaded")

    del downloads[token]

    return StreamingResponse(
        BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=slide.pptx"},
    )


app.mount("/", StaticFiles(directory="static", html=True), name="static")
