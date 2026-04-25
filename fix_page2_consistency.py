"""
Consistency pass on page 2 styles:
1. Fee table intro paragraph: 9pt Evye Small → 10pt body text
2. Table header: 8.5pt → 9pt (match table body)
3. Standardize all table text at 9pt
Result: only 3 sizes in the document — 14pt sections, 10pt body, 9pt tables
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "evye-master.docx"


def fix_table_styles(doc):
    """Standardize all table cell styles to 9pt."""
    print("Standardizing table styles to 9pt...")

    styles_to_fix = {
        "Evye Table Header": 9,      # was 8.5pt
        "Evye Table Body": 9,         # was 9pt (confirm)
        "Evye Table Item Title": 9,   # was 9pt (confirm)
        "Evye Table Total": 9,        # was 9pt (confirm)
        "Evye Bullet Small": 9,       # was 9pt (confirm)
    }

    for name, size in styles_to_fix.items():
        try:
            style = doc.styles[name]
            old_size = style.font.size / 12700 if style.font.size else "inherited"
            style.font.size = Pt(size)
            print(f"  {name}: {old_size}pt → {size}pt")
        except KeyError:
            print(f"  {name}: not found, skipping")


def fix_small_style(doc):
    """Evye Small should be 10pt — same as body. No reason for a separate smaller size."""
    print("Fixing Evye Small style...")
    try:
        small = doc.styles["Evye Small"]
        old_size = small.font.size / 12700 if small.font.size else "inherited"
        small.font.size = Pt(10)
        print(f"  Evye Small: {old_size}pt → 10pt (matches body)")
    except KeyError:
        print(f"  Evye Small: not found")


def fix_fee_table_intro(doc):
    """Change the fee table intro paragraph from Evye Small to Evye Body."""
    print("Fixing fee table intro paragraph style...")
    for p in doc.paragraphs:
        if "Below is a sample fee schedule" in p.text:
            old_style = p.style.name
            p.style = doc.styles["Evye Body"]
            print(f"  '{p.text[:40]}...': {old_style} → Evye Body")
            return
    print("  Fee table intro paragraph not found")


def main():
    print(f"Opening: {TEMPLATE_PATH}")
    doc = Document(str(TEMPLATE_PATH))

    fix_small_style(doc)
    fix_table_styles(doc)
    fix_fee_table_intro(doc)

    doc.save(str(TEMPLATE_PATH))
    print(f"\nSaved: {TEMPLATE_PATH}")

    # Summary
    print("\n=== Final size inventory ===")
    print("  14pt — H2 section headings (Telegraf Bold)")
    print("  10pt — body text, clause text, bullets, KV tables, signatures")
    print("   9pt — fee table cells only")


if __name__ == "__main__":
    main()
