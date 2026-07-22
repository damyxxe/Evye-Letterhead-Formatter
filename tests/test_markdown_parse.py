def parse(md):
    from word_generator import _parse_markdown
    return _parse_markdown(md)


def test_inline_runs():
    from word_generator import _parse_inline
    runs = _parse_inline("plain **bold** *it* `code` [site](https://evye.co)")
    assert {"t": "bold", "b": True} in runs
    assert {"t": "it", "i": True} in runs
    assert {"t": "code", "c": True} in runs
    assert {"t": "site", "link": "https://evye.co"} in runs


def test_inline_underscore_not_greedy_in_words():
    from word_generator import _parse_inline
    runs = _parse_inline("snake_case_name stays")
    assert "".join(r["t"] for r in runs) == "snake_case_name stays"


def test_paragraph_carries_runs():
    blocks = parse("This is **important** text.")
    para = blocks[0]
    assert para["type"] == "paragraph"
    assert para["text"] == "This is important text."
    assert any(r.get("b") and r["t"] == "important" for r in para["runs"])


def test_numbered_list():
    blocks = parse("1. first\n2. second")
    assert [b["type"] for b in blocks] == ["numbered", "numbered"]
    assert blocks[0]["text"] == "first"


def test_todo_items():
    blocks = parse("- [ ] open\n- [x] done")
    assert blocks[0] == {"type": "todo", "checked": False, "text": "open",
                         "runs": blocks[0]["runs"], "level": 1}
    assert blocks[1]["checked"] is True


def test_fenced_code_block():
    blocks = parse("before\n\n```python\nx = 1\ny = 2\n```\n\nafter")
    cb = [b for b in blocks if b["type"] == "code_block"]
    assert len(cb) == 1
    assert cb[0]["text"] == "x = 1\ny = 2"


def test_plain_bullets_still_work():
    blocks = parse("- one\n- two")
    assert all(b["type"] == "bullet" for b in blocks)
