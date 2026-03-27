# Snap to Slide

Turn a photo of handwritten notes, diagrams, or mixed content into a PowerPoint slide — instantly, using Claude Vision.

## Setup

### 1. Install dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Set your Anthropic API key

```bash
export ANTHROPIC\_API\_KEY=sk-ant-...
```

On Windows (PowerShell):
```powershell
$env:ANTHROPIC\_API\_KEY = "sk-ant-..."
```

### 3. Run the server

```bash
uvicorn server:app --reload --port 8000
```

Then open \[http://localhost:8000](http://localhost:8000) in your browser.

## How it works

1\. Upload a photo (JPG, PNG, WEBP, GIF)
2\. Watch live progress as Claude Vision analyzes the image
3\. Download your `.pptx` file

## Slide layouts

Claude picks the best layout for your content:

| Layout | Best for |
|---|---|
| `bullets` | Notes, lists, outlines |
| `two\_column` | Comparisons, two topics |
| `key\_stats` | Numbers, metrics, data |
| `title\_content` | Single topic with explanation |

## Project structure

```
snap-to-slide/
├── server.py          # FastAPI backend
├── requirements.txt
├── static/
│   └── index.html     # Frontend (single file)
└── README.md
```