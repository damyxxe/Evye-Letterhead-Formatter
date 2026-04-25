"""
Fix template resilience and markdown mapping:
1. Set correct 'next paragraph style' so Enter key behaves properly
2. Configure built-in List Bullet / List Number styles to match Evye formatting
3. Unhide Heading 4
4. Map built-in styles to Evye equivalents for markdown paste compatibility
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "evye-master.docx"

HELVETICA = "Helvetica Neue"
FALLBACK = "Arial"


def fix_next_paragraph_styles(doc):
    """Set correct 'next paragraph style' so Enter drops to the right style."""
    print("Fixing next-paragraph-style (Enter key behavior)...")

    mappings = {
        # One-shot styles → drop to body text after Enter
        "Evye Clause": "Evye Subclause",
        "Evye Salutation": "Evye Body",
        "Evye Subtitle": "Evye Body",
        "Evye Signatory Name": "Evye Signatory Title",
        "Evye Signatory Title": "Normal",
        "Evye Cover Title": "Evye Subtitle",
        "Evye Annex Title": "Evye Body",

        # Continuing styles → stay in same style (user keeps typing bullets/subclauses)
        "Evye Subclause": "Evye Subclause",
        "Evye Sub-Subclause": "Evye Sub-Subclause",
        "Evye Bullet": "Evye Bullet",
        "Evye Bullet 2": "Evye Bullet 2",
        "Evye Bullet 3": "Evye Bullet 3",
        "Evye Body": "Evye Body",
        "Evye Body Left": "Evye Body Left",
    }

    for style_name, next_name in mappings.items():
        try:
            style = doc.styles[style_name]
            next_style = doc.styles[next_name]
            style.next_paragraph_style = next_style
            print(f"  {style_name} → Enter → {next_name}")
        except KeyError as e:
            print(f"  {style_name}: style not found ({e})")


def fix_builtin_list_styles(doc):
    """
    Configure Word's built-in List Bullet / List Number styles to match
    Evye bullet formatting. This makes markdown paste work correctly.
    """
    print("Configuring built-in list styles for markdown compatibility...")

    # ── List Bullet (level 1) → matches Evye Bullet ──
    try:
        lb = doc.styles["List Bullet"]
        lb.font.name = HELVETICA
        lb.font.size = Pt(10)
        lb.paragraph_format.left_indent = Mm(10)
        lb.paragraph_format.first_line_indent = Mm(-6)
        lb.paragraph_format.space_after = Pt(2)
        lb.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        lb.paragraph_format.right_indent = Mm(42)
        _clear_tab_stops(lb)
        lb.paragraph_format.tab_stops.add_tab_stop(Mm(10), WD_TAB_ALIGNMENT.LEFT)
        _set_font_xml(lb, HELVETICA, FALLBACK)
        print(f"  List Bullet: → matches Evye Bullet (10mm indent)")
    except KeyError:
        print(f"  List Bullet: not found")

    # ── List Bullet 2 → matches Evye Bullet 2 ──
    try:
        lb2 = doc.styles["List Bullet 2"]
        lb2.font.name = HELVETICA
        lb2.font.size = Pt(10)
        lb2.paragraph_format.left_indent = Mm(18)
        lb2.paragraph_format.first_line_indent = Mm(-6)
        lb2.paragraph_format.space_after = Pt(2)
        lb2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        lb2.paragraph_format.right_indent = Mm(42)
        _clear_tab_stops(lb2)
        lb2.paragraph_format.tab_stops.add_tab_stop(Mm(18), WD_TAB_ALIGNMENT.LEFT)
        _set_font_xml(lb2, HELVETICA, FALLBACK)
        print(f"  List Bullet 2: → matches Evye Bullet 2 (18mm indent)")
    except KeyError:
        print(f"  List Bullet 2: not found")

    # ── List Bullet 3 → matches Evye Bullet 3 ──
    try:
        lb3 = doc.styles["List Bullet 3"]
        lb3.font.name = HELVETICA
        lb3.font.size = Pt(10)
        lb3.paragraph_format.left_indent = Mm(26)
        lb3.paragraph_format.first_line_indent = Mm(-6)
        lb3.paragraph_format.space_after = Pt(2)
        lb3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        lb3.paragraph_format.right_indent = Mm(42)
        _clear_tab_stops(lb3)
        lb3.paragraph_format.tab_stops.add_tab_stop(Mm(26), WD_TAB_ALIGNMENT.LEFT)
        _set_font_xml(lb3, HELVETICA, FALLBACK)
        print(f"  List Bullet 3: → matches Evye Bullet 3 (26mm indent)")
    except KeyError:
        print(f"  List Bullet 3: not found")

    # ── List Number (level 1) → matches Evye Subclause indent ──
    try:
        ln = doc.styles["List Number"]
        ln.font.name = HELVETICA
        ln.font.size = Pt(10)
        ln.paragraph_format.left_indent = Mm(12)
        ln.paragraph_format.first_line_indent = Mm(-12)
        ln.paragraph_format.space_after = Pt(4)
        ln.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ln.paragraph_format.right_indent = Mm(42)
        _clear_tab_stops(ln)
        ln.paragraph_format.tab_stops.add_tab_stop(Mm(12), WD_TAB_ALIGNMENT.LEFT)
        _set_font_xml(ln, HELVETICA, FALLBACK)
        print(f"  List Number: → matches clause indent (12mm tab)")
    except KeyError:
        print(f"  List Number: not found")

    # ── List Number 2 → matches Sub-subclause indent ──
    try:
        ln2 = doc.styles["List Number 2"]
        ln2.font.name = HELVETICA
        ln2.font.size = Pt(10)
        ln2.paragraph_format.left_indent = Mm(22)
        ln2.paragraph_format.first_line_indent = Mm(-10)
        ln2.paragraph_format.space_after = Pt(4)
        ln2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ln2.paragraph_format.right_indent = Mm(42)
        _clear_tab_stops(ln2)
        ln2.paragraph_format.tab_stops.add_tab_stop(Mm(22), WD_TAB_ALIGNMENT.LEFT)
        _set_font_xml(ln2, HELVETICA, FALLBACK)
        print(f"  List Number 2: → matches sub-subclause indent (22mm tab)")
    except KeyError:
        print(f"  List Number 2: not found")

    # ── Quote → styled for Evye brand ──
    try:
        q = doc.styles["Quote"]
        q.font.name = HELVETICA
        q.font.size = Pt(10)
        q.font.italic = True
        q.paragraph_format.left_indent = Mm(10)
        q.paragraph_format.right_indent = Mm(42)
        q.paragraph_format.space_before = Pt(6)
        q.paragraph_format.space_after = Pt(6)
        _set_font_xml(q, HELVETICA, FALLBACK)
        print(f"  Quote: → Helvetica Neue 10pt italic, 10mm indent")
    except KeyError:
        print(f"  Quote: not found")


def fix_visibility(doc):
    """Unhide Heading 4 so it appears in the style picker."""
    print("Fixing style visibility...")
    try:
        h4 = doc.styles["Heading 4"]
        h4.hidden = False
        h4.quick_style = True
        print(f"  Heading 4: unhidden, added to quick styles")
    except KeyError:
        print(f"  Heading 4: not found")


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

    fix_next_paragraph_styles(doc)
    fix_builtin_list_styles(doc)
    fix_visibility(doc)

    doc.save(str(TEMPLATE_PATH))
    print(f"\nSaved: {TEMPLATE_PATH}")

    print("\n=== MARKDOWN → WORD MAPPING ===")
    print("  # Title        → Heading 1 (32pt Telegraf)")
    print("  ## Section     → Heading 2 (14pt Telegraf Bold)")
    print("  ### Subsection → Heading 3 (10pt Helvetica Bold UPPERCASE)")
    print("  Body text      → Normal (10pt Helvetica Neue)")
    print("  - bullet       → List Bullet (10mm indent)")
    print("    - sub        → List Bullet 2 (18mm indent)")
    print("      - sub-sub  → List Bullet 3 (26mm indent)")
    print("  1. numbered    → List Number (12mm tab)")
    print("  > quote        → Quote (10pt italic, 10mm indent)")
    print("  **bold**       → Bold run")
    print("  *italic*       → Italic run")
    print("  | table |      → Table (9pt cells)")


if __name__ == "__main__":
    main()
