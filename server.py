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
from pptx.enum.text import PP_ALIGN
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
        # Make it semi-transparent by using a lighter tone — just use accent
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
    # Soft panel: lighten by mixing primary with bg — use primary at low opacity
    # We'll just use a slightly tinted color
    panel.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
    panel.line.fill.background()
    # Overlay text on top via textbox
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
    }
    builder = builders.get(layout, build_bullets_slide)
    builder(prs, data)

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()


CLAUDE_PROMPT = """You are a PowerPoint slide designer. Analyze the image and extract its content, then design a single professional slide.

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
        # Step 1
        yield f"data: {json.dumps({'step': 1, 'message': 'Sending image to Claude Vision…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.1)

        b64 = base64.standard_b64encode(image_data).decode("utf-8")

        # Step 2
        yield f"data: {json.dumps({'step': 1, 'message': 'Sending image to Claude Vision…', 'status': 'done'})}\n\n"
        yield f"data: {json.dumps({'step': 2, 'message': 'Extracting text and layout…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.1)

        client = anthropic.Anthropic(api_key=api_key)

        loop = asyncio.get_event_loop()
        response = await loop.run_in_executor(
            None,
            lambda: client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=2048,
                messages=[
                    {
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
                            {"type": "text", "text": CLAUDE_PROMPT},
                        ],
                    }
                ],
            ),
        )

        raw_text = response.content[0].text.strip()

        # Step 3
        yield f"data: {json.dumps({'step': 2, 'message': 'Extracting text and layout…', 'status': 'done'})}\n\n"
        yield f"data: {json.dumps({'step': 3, 'message': 'Designing slide structure…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.1)

        # Parse JSON
        if raw_text.startswith("```"):
            lines = raw_text.split("\n")
            raw_text = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])

        slide_data = json.loads(raw_text)

        # Step 4
        yield f"data: {json.dumps({'step': 3, 'message': 'Designing slide structure…', 'status': 'done'})}\n\n"
        yield f"data: {json.dumps({'step': 4, 'message': 'Building PowerPoint file…', 'status': 'active'})}\n\n"
        await asyncio.sleep(0.1)

        pptx_bytes = await loop.run_in_executor(None, lambda: build_pptx(slide_data))

        token = str(uuid.uuid4())
        downloads[token] = pptx_bytes

        # Clean up job
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
