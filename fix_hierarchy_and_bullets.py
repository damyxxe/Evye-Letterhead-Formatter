"""
Fix heading hierarchy, clause tab stops, and bullet alignment
in the user's saved evye-master.docx.
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Mm
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "evye-master.docx"

TELEGRAF = "Telegraf"
HELVETICA = "Helvetica Neue"
FALLBACK = "Arial"


def fix_heading_hierarchy(doc):
    """Fix H2/H3/H4 sizes for clear visual staircase: 32 → 14 → 11 → 10."""
    print("Fixing heading hierarchy...")

    # H2: Telegraf Bold 14pt (was 21pt — too close to H1 at 32pt)
    h2 = doc.styles["Heading 2"]
    h2.font.size = Pt(14)
    h2.font.bold = True
    print(f"  H2: → 14pt Bold")

    # H3: Helvetica Neue Bold 11pt (was 10.5pt Telegraf Bold)
    h3 = doc.styles["Heading 3"]
    h3.font.name = HELVETICA
    h3.font.size = Pt(11)
    h3.font.bold = True
    h3.font.italic = False
    _set_font_xml(h3, HELVETICA, FALLBACK)
    print(f"  H3: → 11pt Helvetica Neue Bold")

    # H4: Helvetica Neue Bold Italic 10pt (was 10pt — keep size, ensure italic)
    h4 = doc.styles["Heading 4"]
    h4.font.name = HELVETICA
    h4.font.size = Pt(10)
    h4.font.bold = True
    h4.font.italic = True
    _set_font_xml(h4, HELVETICA, FALLBACK)
    print(f"  H4: → 10pt Helvetica Neue Bold Italic")


def fix_clause_tab_stops(doc):
    """Widen clause tab stops: 12mm → 18mm, subclauses proportional."""
    print("Fixing clause tab stops...")

    # Evye Clause: tab stop at 18mm
    clause = doc.styles["Evye Clause"]
    _clear_tab_stops(clause)
    clause.paragraph_format.tab_stops.add_tab_stop(Mm(18))
    clause.paragraph_format.left_indent = Mm(0)
    clause.paragraph_format.first_line_indent = Mm(0)
    print(f"  Evye Clause: tab at 18mm")

    # Evye Subclause: left indent 18mm, hanging indent, tab at 30mm
    sub = doc.styles["Evye Subclause"]
    _clear_tab_stops(sub)
    sub.paragraph_format.left_indent = Mm(30)
    sub.paragraph_format.first_line_indent = Mm(-12)
    sub.paragraph_format.tab_stops.add_tab_stop(Mm(30))
    print(f"  Evye Subclause: indent 18mm, tab at 30mm")

    # Evye Sub-Subclause: left indent 30mm, tab at 44mm
    subsub = doc.styles["Evye Sub-Subclause"]
    _clear_tab_stops(subsub)
    subsub.paragraph_format.left_indent = Mm(44)
    subsub.paragraph_format.first_line_indent = Mm(-14)
    subsub.paragraph_format.tab_stops.add_tab_stop(Mm(44))
    print(f"  Evye Sub-Subclause: indent 30mm, tab at 44mm")


def fix_bullet_alignment(doc):
    """Fix bullet indentation to match reference: generous steps, clear nesting."""
    print("Fixing bullet alignment...")

    # Level 1: left indent 10mm, hanging 6mm (bullet at 4mm, text at 10mm)
    b1 = doc.styles["Evye Bullet"]
    _clear_tab_stops(b1)
    b1.paragraph_format.left_indent = Mm(10)
    b1.paragraph_format.first_line_indent = Mm(-6)
    b1.paragraph_format.tab_stops.add_tab_stop(Mm(10))
    print(f"  Evye Bullet: indent 10mm, hang 6mm")

    # Level 2: left indent 18mm, hanging 6mm (bullet at 12mm, text at 18mm)
    b2 = doc.styles["Evye Bullet 2"]
    _clear_tab_stops(b2)
    b2.paragraph_format.left_indent = Mm(18)
    b2.paragraph_format.first_line_indent = Mm(-6)
    b2.paragraph_format.tab_stops.add_tab_stop(Mm(18))
    print(f"  Evye Bullet 2: indent 18mm, hang 6mm")

    # Level 3: left indent 26mm, hanging 6mm (dash at 20mm, text at 26mm)
    b3 = doc.styles["Evye Bullet 3"]
    _clear_tab_stops(b3)
    b3.paragraph_format.left_indent = Mm(26)
    b3.paragraph_format.first_line_indent = Mm(-6)
    b3.paragraph_format.tab_stops.add_tab_stop(Mm(26))
    print(f"  Evye Bullet 3: indent 26mm, hang 6mm")

    # Small bullet (fee tables): tighten proportionally
    bs = doc.styles["Evye Bullet Small"]
    _clear_tab_stops(bs)
    bs.paragraph_format.left_indent = Mm(6)
    bs.paragraph_format.first_line_indent = Mm(-4)
    bs.paragraph_format.tab_stops.add_tab_stop(Mm(6))
    print(f"  Evye Bullet Small: indent 6mm, hang 4mm")


def update_sample_text(doc):
    """Update sample text descriptions to reflect new sizes."""
    for p in doc.paragraphs:
        if "Heading 2 uses Telegraf UltraLight at 20pt" in p.text:
            for r in p.runs:
                r.text = r.text.replace(
                    "Heading 2 uses Telegraf UltraLight at 20pt",
                    "Heading 2 uses Telegraf Bold at 14pt"
                )
        if "Heading 3 uses Telegraf Bold at 10pt" in p.text:
            for r in p.runs:
                r.text = r.text.replace(
                    "Heading 3 uses Telegraf Bold at 10pt",
                    "Heading 3 uses Helvetica Neue Bold at 11pt"
                )
        if "Heading 4 uses Helvetica Neue Bold at 9.5pt" in p.text:
            for r in p.runs:
                r.text = r.text.replace(
                    "Heading 4 uses Helvetica Neue Bold at 9.5pt",
                    "Heading 4 uses Helvetica Neue Bold Italic at 10pt"
                )


# ── Helpers ──────────────────────────────────────────────────

def _set_font_xml(style, ascii_font, cs_font):
    """Set font names at the XML level for full compatibility."""
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
        style.element.append(rpr)
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is not None:
        rpr.remove(rFonts)
    rFonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} '
        f'w:ascii="{ascii_font}" w:hAnsi="{ascii_font}" '
        f'w:cs="{cs_font}" w:eastAsia="{ascii_font}"/>'
    )
    rpr.insert(0, rFonts)


def _clear_tab_stops(style):
    """Remove all existing tab stops from a style."""
    pPr = style.element.find(qn("w:pPr"))
    if pPr is not None:
        tabs = pPr.find(qn("w:tabs"))
        if tabs is not None:
            pPr.remove(tabs)


def main():
    print(f"Opening: {TEMPLATE_PATH}")
    doc = Document(str(TEMPLATE_PATH))

    fix_heading_hierarchy(doc)
    fix_clause_tab_stops(doc)
    fix_bullet_alignment(doc)
    update_sample_text(doc)

    doc.save(str(TEMPLATE_PATH))
    print(f"\nSaved: {TEMPLATE_PATH}")


if __name__ == "__main__":
    main()
