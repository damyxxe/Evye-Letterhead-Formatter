"""
restyle_template.py — Style system overhaul for evye-master.docx

Changes made:
  1. Rename all Evye* custom styles: drop the "Evye" prefix.
     Both w:styleId (the OOXML reference key) and w:name (the display name)
     are updated so they stay in sync. The _style_id() helper in
     word_generator.py strips spaces from the name to produce the styleId,
     so every new name is designed to satisfy: name.replace(" ","") == styleId.
  2. Update ALL cross-references inside styles.xml that use old Evye* IDs
     (basedOn, next, link attributes) to the new IDs.
  3. Add <w:qFormat/> to every renamed style → shows up in the ribbon
     Quick Styles gallery.
  4. Add <w:uiPriority> to every renamed style → controls gallery ordering.
  5. Remove <w:unhideWhenUsed/> from Heading 2-4 → they're always visible.

The Word built-in Heading1–4 styles are NOT renamed; they already have the
correct styleIds for TOC compatibility and Evye brand typography.
"""

import zipfile
import shutil
from pathlib import Path
from lxml import etree

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE     = Path(__file__).resolve().parent.parent
TEMPLATE = HERE / "templates" / "evye-master.docx"
BACKUP   = HERE / "templates" / "evye-master-pre-restyle.docx"
OUTPUT   = TEMPLATE   # overwrite in-place

# ── Style rename map ──────────────────────────────────────────────────────────
# Format: old_styleId → (new_styleId, new_name, uiPriority)
# Constraint: new_name.replace(" ","") == new_styleId  (invariant for _style_id())
RENAME = {
    # Core body
    "EvyeBody":           ("Body",           "Body",            20),
    "EvyeBodyLeft":       ("BodyLeft",        "Body Left",       21),
    "EvyeSalutation":     ("Salutation",      "Salutation",      22),

    # Document title area
    "EvyeSubtitle":       ("DocSubtitle",     "Doc Subtitle",    15),
    "EvyeCoverTitle":     ("CoverTitle",      "Cover Title",     14),

    # Bullets
    "EvyeBullet":         ("Bullet",          "Bullet",          30),
    "EvyeBullet2":        ("Bullet2",         "Bullet 2",        31),
    "EvyeBullet3":        ("Bullet3",         "Bullet 3",        32),
    "EvyeBulletSmall":    ("BulletSmall",     "Bullet Small",    82),

    # Contract clauses
    "EvyeClause":         ("Clause",          "Clause",          40),
    "EvyeSubclause":      ("Subclause",       "Subclause",       41),
    "EvyeSub-Subclause":  ("Sub-Subclause",   "Sub-Subclause",   42),

    # Tables
    "EvyeTableHeader":    ("TableHeader",     "Table Header",    50),
    "EvyeTableBody":      ("TableBody",       "Table Body",      51),
    "EvyeTableTotal":     ("TableTotal",      "Table Total",     52),
    "EvyeTableItemTitle": ("TableItemTitle",  "Table Item Title", 83),
    "EvyeKVLabel":        ("KVLabel",         "KV Label",        53),
    "EvyeKVValue":        ("KVValue",         "KV Value",        54),

    # Signatory
    "EvyeSignatoryName":  ("SignatoryName",   "Signatory Name",  60),
    "EvyeSignatoryTitle": ("SignatoryTitle",  "Signatory Title", 61),

    # Misc
    "EvyeAnnexTitle":     ("AnnexTitle",      "Annex Title",     70),
    "EvyeSmall":          ("Small",           "Small",           80),
    "EvyeTiny":           ("Tiny",            "Tiny",            81),
}

# Mapping: old_styleId → new_styleId  (for cross-reference rewriting)
OLD_TO_NEW_ID = {old: new_id for old, (new_id, _, _) in RENAME.items()}

NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
QN  = lambda tag: f"{{{NS}}}{tag}"


def _set_or_add(parent, tag, attrib):
    """Find-or-create a direct child element and set its attributes."""
    el = parent.find(QN(tag))
    if el is None:
        el = etree.SubElement(parent, QN(tag))
    for k, v in attrib.items():
        el.set(QN(k), v)
    return el


def _ensure_qformat(style_el):
    """Add <w:qFormat/> if absent."""
    if style_el.find(QN("qFormat")) is None:
        # Insert after uiPriority / unhideWhenUsed / rsid / or after name
        style_el.append(etree.Element(QN("qFormat")))


def _set_ui_priority(style_el, priority: int):
    """Set or update <w:uiPriority w:val="N"/>."""
    el = style_el.find(QN("uiPriority"))
    if el is None:
        el = etree.SubElement(style_el, QN("uiPriority"))
    el.set(QN("val"), str(priority))


def _rewrite_ref(style_el, tag, old_to_new):
    """Rewrite a w:val attribute on a child element if it's in old_to_new map."""
    el = style_el.find(QN(tag))
    if el is not None:
        old = el.get(QN("val"), "")
        if old in old_to_new:
            el.set(QN("val"), old_to_new[old])


def transform_styles(xml_bytes: bytes) -> bytes:
    tree = etree.fromstring(xml_bytes)
    styles = tree.findall(QN("style"))

    for style_el in styles:
        sid = style_el.get(QN("styleId"), "")
        if sid not in RENAME:
            continue

        new_sid, new_name, ui_priority = RENAME[sid]

        # 1. Rename styleId attribute
        style_el.set(QN("styleId"), new_sid)

        # 2. Rename w:name val
        name_el = style_el.find(QN("name"))
        if name_el is not None:
            name_el.set(QN("val"), new_name)

        # 3. Add qFormat
        _ensure_qformat(style_el)

        # 4. Set uiPriority
        _set_ui_priority(style_el, ui_priority)

        # 5. Remove <w:unhideWhenUsed/> (we want these visible always)
        uhu = style_el.find(QN("unhideWhenUsed"))
        if uhu is not None:
            style_el.remove(uhu)

    # Second pass: rewrite cross-references (basedOn, next, link)
    # that point to old Evye* IDs.  Must happen after all IDs are renamed.
    for style_el in tree.findall(QN("style")):
        _rewrite_ref(style_el, "basedOn", OLD_TO_NEW_ID)
        _rewrite_ref(style_el, "next",    OLD_TO_NEW_ID)
        _rewrite_ref(style_el, "link",    OLD_TO_NEW_ID)

    return etree.tostring(tree, xml_declaration=True,
                          encoding="UTF-8", standalone=True)


def rebuild_docx():
    # Back up original
    if not BACKUP.exists():
        shutil.copy2(TEMPLATE, BACKUP)
        print(f"Backed up to {BACKUP.name}")
    else:
        print(f"Backup already exists: {BACKUP.name}")

    # Read all files from the template zip
    file_map = {}
    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        for name in zin.namelist():
            file_map[name] = zin.read(name)

    # Transform styles.xml
    original  = file_map["word/styles.xml"]
    modified  = transform_styles(original)
    file_map["word/styles.xml"] = modified

    # Write new zip
    tmp = TEMPLATE.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_map.items():
            zout.writestr(name, data)

    tmp.replace(OUTPUT)
    print(f"Template updated: {OUTPUT}")


def verify():
    """Quick sanity check: confirm old IDs are gone, new ones exist."""
    with zipfile.ZipFile(TEMPLATE, "r") as z:
        raw = z.read("word/styles.xml")
    tree = etree.fromstring(raw)
    ids = {s.get(QN("styleId"), "") for s in tree.findall(QN("style"))}

    missing_new = [new for (new, _, _) in RENAME.values() if new not in ids]
    stale_old   = [old for old in RENAME if old in ids]

    if stale_old:
        print(f"  ✗ Stale old IDs still present: {stale_old}")
    else:
        print("  ✓ All old Evye* IDs removed")

    if missing_new:
        print(f"  ✗ Missing new IDs: {missing_new}")
    else:
        print("  ✓ All new IDs present")

    # Verify qFormat on new styles
    no_qfmt = []
    for s in tree.findall(QN("style")):
        sid = s.get(QN("styleId"), "")
        if sid in {new for (new, _, _) in RENAME.values()}:
            if s.find(QN("qFormat")) is None:
                no_qfmt.append(sid)
    if no_qfmt:
        print(f"  ✗ Missing qFormat on: {no_qfmt}")
    else:
        print("  ✓ qFormat present on all renamed styles")


if __name__ == "__main__":
    rebuild_docx()
    print("Verifying…")
    verify()
    print("Done.")
