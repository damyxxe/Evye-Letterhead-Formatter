import json
from pathlib import Path
from conftest import doc_paragraphs, doc_tables, hyperlinks

FIXTURE = Path(__file__).parent / "fixtures" / "notion_all_blocks.json"


def test_every_block_type_lands_in_the_document(gen):
    from word_generator import generate_from_notion_blocks
    fx = json.loads(FIXTURE.read_text())
    path = gen(generate_from_notion_blocks, fx["blocks"],
               page_properties=fx["page_properties"])

    paras = doc_paragraphs(path)
    texts = [t for _, t in paras]
    joined = "\n".join(texts)

    # Title split on em-dash
    assert ("Heading 1", "Fixture Doc") in paras
    # Inline formatting survived
    assert any("Intro with bold and a link" == t for t in texts)
    assert ("link", "https://evye.co") in hyperlinks(path)
    # Heading re-levelling
    assert ("Heading 2", "H1 Section") in paras
    assert ("Heading 3", "H2 Sub") in paras
    assert ("Heading 4", "H3 Subsub") in paras
    # Lists
    assert any(t == "•\tbullet one" for s, t in paras if s == "Bullet")
    assert any("nested bullet" in t for s, t in paras if s == "Bullet 2")
    assert any(t.startswith("1.\tstep one") for t in texts)
    assert any(t.startswith("2.\tstep two") for t in texts)
    assert any(t.startswith("☐\topen task") for t in texts)
    assert any(t.startswith("☑\tdone task") for t in texts)
    # Toggle contents
    assert "Toggle header" in joined and "hidden detail" in joined
    # Quote/callout as asides
    assert any(t == "a wise quote" for s, t in paras if s == "Body Left")
    assert any(t == "important note" for s, t in paras if s == "Body Left")
    # Code + equation
    assert "print('hi')" in joined
    assert "E = mc^2" in joined
    # Bookmark link + child page placeholder + unknown-type placeholder
    assert ("https://anthropic.com", "https://anthropic.com") in hyperlinks(path)
    assert "[Sub-page: Appendix Page]" in joined
    assert "[Unsupported block: mystery_block_type]" in joined
    # Tables: fee table (3-col) and generic (4-col)
    tables = doc_tables(path)
    assert len(tables) == 2
    assert tables[0][0] == ["Item", "Description", "Fee"]
    assert tables[1][0] == ["W", "X", "Y", "Z"]
    # Sign-off promotion: "Warm regards"/"Test Signer" removed from body,
    # signer rendered in signature block
    assert "Warm regards" not in joined
    assert any(t == "Test Signer" for s, t in paras if s == "Signatory Name")
