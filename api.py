"""
Evye Document Generator — HTTP API
====================================
Minimal FastAPI wrapper around word_generator.py.
Deployed behind OpenLiteSpeed at evye.co/docs/
or locally on Alfred at http://host.docker.internal:8091.

Endpoints:
  GET  /              → Web form (index.html)
  POST /api/generate  → Generate .docx (or .zip bundle), return download URL
  GET  /api/health    → Health check
  GET  /dl/{file}     → Download generated .docx or .zip files

bundle param (POST /api/generate):
  bundle=true  (default) → ZIP bundle with .docx + fonts (web form path)
  bundle=false           → .docx only (bot path — no ZIP needed)
"""

import os
import zipfile
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from word_generator import generate_from_markdown, generate_from_notion_blocks

APP_DIR    = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "output"
WEB_DIR    = APP_DIR / "web"
FONTS_DIR  = APP_DIR / "fonts"
OUTPUT_DIR.mkdir(exist_ok=True)

# Base URL for download links — configurable via env var.
# evye.co (behind OLS proxy):    BASE_URL=/docs
# Alfred (direct, from Docker):  BASE_URL=http://host.docker.internal:8091
BASE_URL = os.getenv("BASE_URL", "/docs")

# Font files to include in bundles
BUNDLE_FONTS = ["Telegraf-Regular.ttf", "Telegraf-Bold.ttf"]

app = FastAPI(title="Evye Document Generator", docs_url=None, redoc_url=None)


# ── Models ────────────────────────────────────────────────────────────────────

class GenerateRequest(BaseModel):
    markdown: str
    metadata: dict = {}
    bundle: bool = True   # False → return .docx only (bot path); True → ZIP + fonts (web form)

class GenerateFromBlocksRequest(BaseModel):
    blocks: list                    # Notion API block objects (tables must have _rows pre-fetched)
    page_properties: dict = {}      # Full GET /v1/pages/{id} response (for title extraction)
    metadata: dict = {}             # Optional metadata overrides
    bundle: bool = False            # Bot default is False (.docx only); True = ZIP+fonts

class GenerateResponse(BaseModel):
    status: str
    download_url: str
    filename: str


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/docs/", response_class=HTMLResponse)
@app.get("/docs", response_class=HTMLResponse)
def serve_form():
    index = WEB_DIR / "index.html"
    if not index.exists():
        raise HTTPException(404, "Web form not found")
    return index.read_text()


@app.post("/docs/api/generate", response_model=GenerateResponse)
def generate(req: GenerateRequest):
    if not req.markdown.strip():
        raise HTTPException(400, "markdown field is required and cannot be empty")

    try:
        path = generate_from_markdown(
            req.markdown,
            req.metadata or {},
            filename=None,   # auto-slug: {date}_{doc_type}_{title-slug}
        )
    except Exception as e:
        raise HTTPException(500, f"Generation failed: {e}")

    docx_path = Path(path)
    docx_name = docx_path.name

    if not req.bundle:
        # Bot path — return the .docx directly (no ZIP needed)
        return GenerateResponse(
            status="ok",
            download_url=f"{BASE_URL}/dl/{docx_name}",
            filename=docx_name,
        )

    # Web form path — create a ZIP bundle: .docx + font files + README
    zip_name = docx_path.stem + ".zip"
    zip_path = OUTPUT_DIR / zip_name
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(docx_path, docx_name)
        for font_file in BUNDLE_FONTS:
            font_path = FONTS_DIR / font_file
            if font_path.exists():
                zf.write(font_path, f"fonts/{font_file}")
        zf.writestr("fonts/INSTALL.txt",
            "Evye LLP — Font Installation\n"
            "=============================\n\n"
            "For the document to display correctly, install these fonts:\n\n"
            "  1. Double-click each .otf file in this folder\n"
            "  2. Click 'Install Font' when prompted\n"
            "  3. Open the .docx file — it will now render with the correct fonts\n\n"
            "You only need to install the fonts once per computer.\n"
        )

    return GenerateResponse(
        status="ok",
        download_url=f"{BASE_URL}/dl/{zip_name}",
        filename=zip_name,
    )


@app.post("/docs/api/generate-from-blocks", response_model=GenerateResponse)
def generate_from_blocks(req: GenerateFromBlocksRequest):
    """
    Generate a .docx from raw Notion API block objects.

    The bot path: n8n fetches Notion blocks (including table rows) and sends
    them here. All smart preprocessing (preamble strip, recipient extraction,
    signature strip, table detection) happens in Python — no JS conversion
    needed in n8n.

    bundle=false (default) → .docx URL only
    bundle=true            → ZIP with .docx + fonts (web form path)
    """
    if not req.blocks:
        raise HTTPException(400, "blocks field is required and cannot be empty")

    try:
        path = generate_from_notion_blocks(
            req.blocks,
            page_properties=req.page_properties or {},
            metadata_overrides=req.metadata or {},
            filename=None,   # auto-slug: {date}_{doc_type}_{title-slug}
        )
    except Exception as e:
        raise HTTPException(500, f"Generation failed: {e}")

    docx_path = Path(path)
    docx_name = docx_path.name

    if not req.bundle:
        return GenerateResponse(
            status="ok",
            download_url=f"{BASE_URL}/dl/{docx_name}",
            filename=docx_name,
        )

    # Bundle path (same as /api/generate with bundle=True)
    zip_name = docx_path.stem + ".zip"
    zip_path = OUTPUT_DIR / zip_name
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(docx_path, docx_name)
        for font_file in BUNDLE_FONTS:
            font_path = FONTS_DIR / font_file
            if font_path.exists():
                zf.write(font_path, f"fonts/{font_file}")
        zf.writestr("fonts/INSTALL.txt",
            "Evye LLP — Font Installation\n"
            "=============================\n\n"
            "For the document to display correctly, install these fonts:\n\n"
            "  1. Double-click each .otf file in this folder\n"
            "  2. Click 'Install Font' when prompted\n"
            "  3. Open the .docx file — it will now render with the correct fonts\n\n"
            "You only need to install the fonts once per computer.\n"
        )
    return GenerateResponse(
        status="ok",
        download_url=f"{BASE_URL}/dl/{zip_name}",
        filename=zip_name,
    )


@app.get("/docs/api/health")
def health():
    return {"status": "ok"}


@app.get("/docs/dl/{filename}")
def download(filename: str):
    if "/" in filename or ".." in filename:
        raise HTTPException(400, "Invalid filename")
    if not (filename.endswith(".docx") or filename.endswith(".zip")):
        raise HTTPException(400, "Invalid file type")
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(404, "File not found or expired")
    media = "application/zip" if filename.endswith(".zip") else \
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path, media_type=media, filename=filename)


# ── Static assets (fonts for the web form) ────────────────────────────────────

if (APP_DIR / "fonts").exists():
    app.mount("/docs/fonts", StaticFiles(directory=str(APP_DIR / "fonts")), name="fonts")
