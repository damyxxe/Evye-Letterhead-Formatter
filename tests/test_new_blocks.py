from conftest import doc_paragraphs, doc_tables, image_count

# 1x1 transparent PNG as a data URL — lets the image test run offline
PNG_1PX = ("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAf"
           "FcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")


def _make(gen, body):
    from word_generator import generate_docx
    return gen(generate_docx, "custom", {"doc_title": "", "body": body})


def test_numbered_list_renders_with_numbers(gen):
    path = _make(gen, [
        {"type": "numbered", "text": "first", "level": 1},
        {"type": "numbered", "text": "second", "level": 1},
        {"type": "numbered", "text": "sub one", "level": 2},
        {"type": "numbered", "text": "third", "level": 1},
    ])
    texts = [t for _, t in doc_paragraphs(path)]
    assert any(t.startswith("1.\tfirst") for t in texts)
    assert any(t.startswith("2.\tsecond") for t in texts)
    assert any(t.startswith("1.\tsub one") for t in texts)   # sub-list restarts
    assert any(t.startswith("3.\tthird") for t in texts)     # parent continues


def test_numbered_resets_after_interruption(gen):
    path = _make(gen, [
        {"type": "numbered", "text": "one", "level": 1},
        {"type": "paragraph", "text": "interlude"},
        {"type": "numbered", "text": "fresh", "level": 1},
    ])
    texts = [t for _, t in doc_paragraphs(path)]
    assert any(t.startswith("1.\tfresh") for t in texts)


def test_todo_checkboxes(gen):
    path = _make(gen, [
        {"type": "todo", "checked": True, "text": "done thing", "level": 1},
        {"type": "todo", "checked": False, "text": "open thing", "level": 1},
    ])
    texts = [t for _, t in doc_paragraphs(path)]
    assert any(t.startswith("☑\tdone thing") for t in texts)
    assert any(t.startswith("☐\topen thing") for t in texts)


def test_code_block(gen):
    path = _make(gen, [{"type": "code_block", "text": "def f():\n    return 1"}])
    paras = doc_paragraphs(path)
    idx = next(i for i, (s, t) in enumerate(paras) if "def f():" in t)
    assert paras[idx][0] == "Body Left"
    from conftest import para_runs
    assert all(r["font"] == "Courier New" for r in para_runs(path, idx) if r["text"].strip())


def test_image_embedded(gen):
    path = _make(gen, [{"type": "image", "url": PNG_1PX, "caption": "tiny"}])
    assert image_count(path) == 1
    assert any(t == "tiny" for _, t in doc_paragraphs(path))


def test_image_fallback_on_bad_url(gen):
    path = _make(gen, [{"type": "image", "url": "https://invalid.invalid/x.png", "caption": ""}])
    assert image_count(path) == 0
    assert any("[Image unavailable" in t for _, t in doc_paragraphs(path))


def test_generic_table_four_columns(gen):
    path = _make(gen, [{
        "type": "generic_table",
        "headers": ["A", "B", "C", "D"],
        "rows": [["1", "2", "3", "4"], ["5", "6", "7", "8"]],
    }])
    tables = doc_tables(path)
    assert len(tables) == 1
    assert tables[0][0] == ["A", "B", "C", "D"]
    assert tables[0][2] == ["5", "6", "7", "8"]


def test_generic_table_one_column(gen):
    path = _make(gen, [{"type": "generic_table", "headers": None,
                        "rows": [["only"], ["cells"]]}])
    tables = doc_tables(path)
    assert tables[0] == [["only"], ["cells"]]
