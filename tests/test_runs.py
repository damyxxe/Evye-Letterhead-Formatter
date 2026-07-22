from conftest import doc_paragraphs, para_runs, hyperlinks


def _make(gen, body):
    from word_generator import generate_docx
    return gen(generate_docx, "custom", {"doc_title": "", "body": body})


def test_bold_italic_runs(gen):
    path = _make(gen, [{
        "type": "paragraph",
        "text": "plain bold italic",
        "runs": [
            {"t": "plain "},
            {"t": "bold", "b": True},
            {"t": " "},
            {"t": "italic", "i": True},
        ],
    }])
    # paragraph 0 is the required empty body para? No — custom builds title block
    # then body; with empty title the first content para holds our runs.
    idx = next(i for i, (s, t) in enumerate(doc_paragraphs(path)) if "plain" in t)
    runs = para_runs(path, idx)
    assert [r["text"] for r in runs] == ["plain ", "bold", " ", "italic"]
    assert runs[1]["bold"] and not runs[1]["italic"]
    assert runs[3]["italic"] and not runs[3]["bold"]


def test_underline_strike_code(gen):
    path = _make(gen, [{
        "type": "paragraph",
        "text": "u s c",
        "runs": [
            {"t": "u", "u": True},
            {"t": "s", "s": True},
            {"t": "c", "c": True},
        ],
    }])
    idx = next(i for i, (s, t) in enumerate(doc_paragraphs(path)) if t == "usc")
    runs = para_runs(path, idx)
    assert runs[0]["underline"]
    assert runs[1]["strike"]
    assert runs[2]["font"] == "Courier New"


def test_hyperlink(gen):
    path = _make(gen, [{
        "type": "paragraph",
        "text": "see the site",
        "runs": [{"t": "see "}, {"t": "the site", "link": "https://evye.co"}],
    }])
    links = hyperlinks(path)
    assert ("the site", "https://evye.co") in links


def test_plain_string_still_works(gen):
    path = _make(gen, [{"type": "paragraph", "text": "just a string"}])
    assert any(t == "just a string" for _, t in doc_paragraphs(path))
