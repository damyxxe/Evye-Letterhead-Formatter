"""
Update evye-master.docx styles in-place:
1. Scale all font sizes proportionately (9.5pt → 10pt base)
2. Add DATE autofield to the header text box after www.evye.co
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Mm
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree
from copy import deepcopy

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "evye-master.docx"

HELVETICA = "Helvetica Neue"
FALLBACK = "Arial"

# ── Scale factor: 10pt / 9.5pt ──────────────────────────────
SCALE = 10.0 / 9.5

# Size mappings: old_pt → new_pt (manually rounded for clean values)
SIZE_MAP = {
    9.5: 10.0,    # body text → new base
    8.0: 8.5,     # subtitle, table header
    8.5: 9.0,     # small text, table body, item descriptions
    7.0: 7.5,     # tiny text
    6.5: 7.0,     # footer confidential
    10.0: 10.5,   # clause headings, H3 → slightly larger
    20.0: 21.0,   # H2, annex title
    22.0: 22.0,   # H1 style (user already set run to 32pt, keep style at 22)
    28.0: 29.0,   # cover title
}


def scale_size(old_pt):
    """Map old font size to new scaled size."""
    if old_pt in SIZE_MAP:
        return SIZE_MAP[old_pt]
    # For unmapped sizes, apply ratio
    return round(old_pt * SCALE, 1)


def update_styles(doc):
    """Update all Evye and Heading styles to new proportional sizes."""
    for style in doc.styles:
        name = style.name
        if not (name.startswith("Evye") or name.startswith("Heading")):
            continue

        if style.font.size:
            old_pt = style.font.size / 12700  # EMU to pt
            new_pt = scale_size(old_pt)
            if new_pt != old_pt:
                style.font.size = Pt(new_pt)
                print(f"  {name}: {old_pt:.1f}pt → {new_pt:.1f}pt")

    # Update Normal style
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10)
    print(f"  Normal: 9.5pt → 10.0pt")


def update_footer_sizes(doc):
    """Scale footer text sizes."""
    footer = doc.sections[0].footer
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for p in footer._element.iter(qn('w:r')):
        sz = p.find(qn('w:rPr'))
        if sz is not None:
            sz_el = sz.find(qn('w:sz'))
            if sz_el is not None:
                old_half_pt = int(sz_el.get(qn('w:val')))
                old_pt = old_half_pt / 2
                new_pt = scale_size(old_pt)
                new_half_pt = int(new_pt * 2)
                if new_half_pt != old_half_pt:
                    sz_el.set(qn('w:val'), str(new_half_pt))
                    # Also update szCs if present
                    szCs = sz.find(qn('w:szCs'))
                    if szCs is not None:
                        szCs.set(qn('w:val'), str(new_half_pt))


def add_date_field_to_textbox(doc):
    """Add a DATE autofield paragraph in the header text box, after www.evye.co."""
    header = doc.sections[0].header
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    }

    # Find the text box content
    txbx_content = header._element.find('.//wps:txbx/w:txbxContent', ns)
    if txbx_content is None:
        print("  WARNING: No text box found in header!")
        return

    # Find the paragraph containing "www.evye.co"
    target_para = None
    for p in txbx_content.findall('w:p', ns):
        texts = p.findall('.//w:t', ns)
        full_text = ''.join(t.text or '' for t in texts)
        if 'www.evye.co' in full_text:
            target_para = p
            break

    if target_para is None:
        print("  WARNING: Could not find 'www.evye.co' paragraph in text box!")
        return

    # Get the rPr (run properties) from the www.evye.co paragraph to match formatting
    existing_rpr = target_para.find('.//w:rPr', ns)
    rpr_xml = etree.tostring(existing_rpr).decode() if existing_rpr is not None else ''

    # Build the DATE field paragraph — matches the style of the other text box paragraphs
    # Using "d MMMM yyyy" format to match Evye date style (e.g., "11 April 2026")
    date_para = parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'  <w:pPr>'
        f'    <w:spacing w:line="336" w:lineRule="auto"/>'
        f'    <w:ind w:right="0"/>'
        f'    <w:jc w:val="right"/>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'  </w:pPr>'
        f'  <w:r>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'    <w:fldChar w:fldCharType="begin"/>'
        f'  </w:r>'
        f'  <w:r>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'    <w:instrText xml:space="preserve"> DATE \\@ "d MMMM yyyy" </w:instrText>'
        f'  </w:r>'
        f'  <w:r>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'    <w:fldChar w:fldCharType="separate"/>'
        f'  </w:r>'
        f'  <w:r>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'    <w:t>11 April 2026</w:t>'
        f'  </w:r>'
        f'  <w:r>'
        f'    <w:rPr>'
        f'      <w:rFonts w:ascii="{HELVETICA}" w:hAnsi="{HELVETICA}" w:cs="{FALLBACK}"/>'
        f'      <w:sz w:val="16"/>'
        f'      <w:szCs w:val="16"/>'
        f'    </w:rPr>'
        f'    <w:fldChar w:fldCharType="end"/>'
        f'  </w:r>'
        f'</w:p>'
    )

    # Insert after the www.evye.co paragraph
    target_para.addnext(date_para)
    print("  Added DATE autofield after www.evye.co")


def main():
    print(f"Opening: {TEMPLATE_PATH}")
    doc = Document(str(TEMPLATE_PATH))

    print("\nScaling styles (9.5pt → 10pt base):")
    update_styles(doc)

    print("\nScaling footer text:")
    update_footer_sizes(doc)

    print("\nAdding date field to header text box:")
    add_date_field_to_textbox(doc)

    # Save
    doc.save(str(TEMPLATE_PATH))
    print(f"\nSaved: {TEMPLATE_PATH}")


if __name__ == "__main__":
    main()
