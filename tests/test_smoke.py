from conftest import doc_paragraphs


def test_letter_smoke(gen):
    from word_generator import generate_docx
    path = gen(generate_docx, "letter", {
        "doc_title": "Smoke Test",
        "recipient_name": "Jane Tan",
        "body": [
            {"type": "paragraph", "text": "Hello world."},
            {"type": "heading2", "text": "Section"},
            {"type": "bullet", "text": "Point one", "level": 1},
        ],
        "signatory_name": "Michael Ryan Chan",
        "signatory_title": "Managing Partner",
    })
    paras = doc_paragraphs(path)
    styles = [s for s, _ in paras]
    texts = [t for _, t in paras]
    assert "Heading 1" in styles          # doc title
    assert any("Dear" in t and "Jane Tan" in t for t in texts)
    assert any(t == "Hello world." for t in texts)
    assert ("Heading 2", "Section") in paras
    assert any("Point one" in t for s, t in paras if s == "Bullet")
    assert any("Evye LLP" == t for t in texts)


def test_version_exists():
    import word_generator
    assert word_generator.__version__.startswith("1.2")
