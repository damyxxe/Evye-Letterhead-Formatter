"""
Fix clause indentation and consolidate heading styles:
1. Tier 1 & 2 clauses flush left, tighter tab stops
2. H3 matches clause heading style (bold uppercase 10pt)
3. Fewer distinct font sizes — simplify the visual system
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Mm
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_TAB_ALIGNMENT

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "evye-master.docx"

HELVETICA = "Helvetica Neue"
TELEGRAF = "Telegraf"
FALLBACK = "Arial"


def fix_styles(doc):
    """
    Consolidate to 3 body sizes: 14pt (H2), 10pt (everything else), small (9pt tables).
    H3 = Clause heading = Bold Uppercase 10pt.
    """
    print("Consolidating heading styles...")

    # H2: Telegraf Bold 14pt — section breaks (keep as-is)
    h2 = doc.styles["Heading 2"]
    h2.font.size = Pt(14)
    h2.font.bold = True
    print(f"  H2: 14pt Telegraf Bold (unchanged)")

    # H3: Helvetica Neue Bold 10pt UPPERCASE — matches clause headings
    h3 = doc.styles["Heading 3"]
    h3.font.name = HELVETICA
    h3.font.size = Pt(10)
    h3.font.bold = True
    h3.font.italic = False
    h3.font.all_caps = True
    _set_font_xml(h3, HELVETICA, FALLBACK)
    print(f"  H3: 10pt Helvetica Neue Bold UPPERCASE")

    # H4: Helvetica Neue Bold 10pt (same size, not uppercase — differentiated by context)
    h4 = doc.styles["Heading 4"]
    h4.font.name = HELVETICA
    h4.font.size = Pt(10)
    h4.font.bold = True
    h4.font.italic = True
    h4.font.all_caps = False
    _set_font_xml(h4, HELVETICA, FALLBACK)
    print(f"  H4: 10pt Helvetica Neue Bold Italic")

    # Clause heading: 10pt Bold UPPERCASE — same visual as H3
    clause = doc.styles["Evye Clause"]
    clause.font.size = Pt(10)
    clause.font.bold = True
    print(f"  Evye Clause: 10pt Bold (matches H3)")


def fix_clause_indentation(doc):
    """
    Tier 1 (1.) and Tier 2 (1.1.) flush left with same tab stop.
    Only Tier 3 (1.2.1.) gets indented.
    Tab stop at 12mm — tighter than before.
    """
    print("Fixing clause indentation...")

    TAB1 = Mm(12)  # Text start for tier 1 & 2
    TAB2 = Mm(22)  # Text start for tier 3

    # Tier 1 — Clause heading: "1." at 0, text at 12mm
    clause = doc.styles["Evye Clause"]
    _clear_tab_stops(clause)
    clause.paragraph_format.left_indent = Mm(0)
    clause.paragraph_format.first_line_indent = Mm(0)
    clause.paragraph_format.tab_stops.add_tab_stop(TAB1, WD_TAB_ALIGNMENT.LEFT)
    print(f"  Clause: left=0, tab=12mm")

    # Tier 2 — Subclause: "1.1." at 0, text at 12mm (flush with tier 1 text)
    sub = doc.styles["Evye Subclause"]
    _clear_tab_stops(sub)
    sub.paragraph_format.left_indent = TAB1
    sub.paragraph_format.first_line_indent = Mm(-12)
    sub.paragraph_format.tab_stops.add_tab_stop(TAB1, WD_TAB_ALIGNMENT.LEFT)
    print(f"  Subclause: left=12mm, hang=-12mm (number at 0, text at 12mm)")

    # Tier 3 — Sub-subclause: "1.2.1." at 12mm, text at 22mm
    subsub = doc.styles["Evye Sub-Subclause"]
    _clear_tab_stops(subsub)
    subsub.paragraph_format.left_indent = TAB2
    subsub.paragraph_format.first_line_indent = Mm(-10)
    subsub.paragraph_format.tab_stops.add_tab_stop(TAB2, WD_TAB_ALIGNMENT.LEFT)
    print(f"  Sub-subclause: left=22mm, hang=-10mm (number at 12mm, text at 22mm)")


def update_sample_text(doc):
    """Update descriptions to reflect new styling."""
    replacements = {
        "Heading 2 uses Telegraf Bold at 14pt": "Heading 2 uses Telegraf Bold at 14pt",
        "Heading 3 uses Helvetica Neue Bold at 11pt for subsection titles":
            "Heading 3 uses Helvetica Neue Bold at 10pt, uppercase — matches clause headings",
        "Heading 4 uses Helvetica Neue Bold Italic at 10pt for the finest heading level":
            "Heading 4 uses Helvetica Neue Bold Italic at 10pt for the finest heading level",
    }
    for p in doc.paragraphs:
        for old, new in replacements.items():
            if old in p.text:
                for r in p.runs:
                    if old in r.text:
                        r.text = r.text.replace(old, new)


# ── Helpers ──────────────────────────────────────────────────

def _set_font_xml(style, ascii_font, cs_font):
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
    pPr = style.element.find(qn("w:pPr"))
    if pPr is not None:
        tabs = pPr.find(qn("w:tabs"))
        if tabs is not None:
            pPr.remove(tabs)


def main():
    print(f"Opening: {TEMPLATE_PATH}")
    doc = Document(str(TEMPLATE_PATH))

    fix_styles(doc)
    fix_clause_indentation(doc)
    update_sample_text(doc)

    doc.save(str(TEMPLATE_PATH))
    print(f"\nSaved: {TEMPLATE_PATH}")


if __name__ == "__main__":
    main()
