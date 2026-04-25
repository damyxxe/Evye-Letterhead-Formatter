"""
Evye LLP — Document PDF Generator
Renders Jinja2 HTML templates into branded PDFs using WeasyPrint.
"""

import os
import sys
from pathlib import Path

from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from weasyprint.text.fonts import FontConfiguration


# Project root (directory containing this file)
PROJECT_ROOT = Path(__file__).resolve().parent

# Paths
TEMPLATES_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"
FONTS_DIR = PROJECT_ROOT / "fonts"
ASSETS_DIR = PROJECT_ROOT / "assets"
OUTPUT_DIR = PROJECT_ROOT / "output"

# Jinja2 environment
_env = Environment(
    loader=FileSystemLoader(str(TEMPLATES_DIR)),
    autoescape=False,
)


def generate_pdf(template_name: str, data: dict, output_path: str | None = None) -> str:
    """
    Render a Jinja2 template with the given data and produce a PDF.

    Args:
        template_name: Template filename (e.g., 'contract.html', 'quotation.html', 'letter.html')
        data: Dict of template variables
        output_path: Output PDF path. If None, writes to output/<template_name>.pdf

    Returns:
        Absolute path to the generated PDF.
    """
    # Resolve output path
    if output_path is None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        stem = Path(template_name).stem
        output_path = str(OUTPUT_DIR / f"{stem}.pdf")

    output_path = str(Path(output_path).resolve())

    # Inject asset paths (resolved as file:// URIs for WeasyPrint)
    data.setdefault("css_path", (STATIC_DIR / "letterhead.css").as_uri())
    data.setdefault("logo_path", (ASSETS_DIR / "evye-logo.png").as_uri())

    # Render HTML
    template = _env.get_template(template_name)
    html_string = template.render(**data)

    # Font configuration
    font_config = FontConfiguration()

    # Build CSS with absolute font paths for WeasyPrint
    font_css = _build_font_css()

    # Generate PDF
    html = HTML(string=html_string, base_url=str(PROJECT_ROOT))
    html.write_pdf(
        output_path,
        stylesheets=[],
        font_config=font_config,
    )

    print(f"PDF generated: {output_path}")
    return output_path


def generate_html(template_name: str, data: dict, output_path: str | None = None) -> str:
    """
    Render a Jinja2 template to an HTML file (for browser preview).

    Args:
        template_name: Template filename
        data: Dict of template variables
        output_path: Output HTML path. If None, writes to output/<template_name>.html

    Returns:
        Absolute path to the generated HTML.
    """
    if output_path is None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        stem = Path(template_name).stem
        output_path = str(OUTPUT_DIR / f"{stem}.html")

    output_path = str(Path(output_path).resolve())

    # For HTML preview, use relative paths
    data.setdefault("css_path", os.path.relpath(STATIC_DIR / "letterhead.css", Path(output_path).parent))
    data.setdefault("logo_path", os.path.relpath(ASSETS_DIR / "evye-logo.png", Path(output_path).parent))

    template = _env.get_template(template_name)
    html_string = template.render(**data)

    Path(output_path).write_text(html_string, encoding="utf-8")
    print(f"HTML generated: {output_path}")
    return output_path


def _build_font_css() -> str:
    """Build @font-face CSS with absolute file paths for WeasyPrint."""
    weights = {
        "UltraLight": 200,
        "Regular": 400,
        "Bold": 700,
        "UltraBold": 900,
    }
    rules = []
    for name, weight in weights.items():
        otf = FONTS_DIR / f"Telegraf-{name}.otf"
        woff2 = FONTS_DIR / f"Telegraf-{name}.woff2"
        src_parts = []
        if woff2.exists():
            src_parts.append(f"url('{woff2.as_uri()}') format('woff2')")
        if otf.exists():
            src_parts.append(f"url('{otf.as_uri()}') format('opentype')")
        if src_parts:
            rules.append(f"""@font-face {{
    font-family: 'Telegraf';
    src: {', '.join(src_parts)};
    font-weight: {weight};
    font-style: normal;
}}""")
    return "\n".join(rules)


if __name__ == "__main__":
    # Quick test: generate a simple letter
    test_data = {
        "doc_title": "Test Letter",
        "date": "29 March 2026",
        "recipient_name": "John Doe",
        "recipient_salutation": "Mr",
        "subject": "Test Document Generation",
        "body": "<p>This is a test document to verify the PDF generation pipeline is working correctly.</p>",
        "signatory_name": "Michael Ryan Chan",
        "signatory_title": "Managing Partner",
    }
    generate_pdf("letter.html", test_data)
