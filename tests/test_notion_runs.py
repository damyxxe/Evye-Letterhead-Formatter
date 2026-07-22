def test_extract_notion_runs():
    from word_generator import _extract_notion_runs
    rt = [
        {"plain_text": "Hello ", "annotations": {"bold": False, "italic": False,
         "underline": False, "strikethrough": False, "code": False}, "href": None},
        {"plain_text": "bold", "annotations": {"bold": True, "italic": False,
         "underline": False, "strikethrough": False, "code": False}, "href": None},
        {"plain_text": "link", "annotations": {"bold": False, "italic": True,
         "underline": False, "strikethrough": False, "code": False},
         "href": "https://evye.co"},
    ]
    runs = _extract_notion_runs(rt)
    assert runs == [
        {"t": "Hello ", "b": False, "i": False, "u": False, "s": False, "c": False, "link": None},
        {"t": "bold", "b": True, "i": False, "u": False, "s": False, "c": False, "link": None},
        {"t": "link", "b": False, "i": True, "u": False, "s": False, "c": False, "link": "https://evye.co"},
    ]


def test_extract_notion_runs_empty_and_missing_annotations():
    from word_generator import _extract_notion_runs
    assert _extract_notion_runs([]) == []
    assert _extract_notion_runs(None) == []
    runs = _extract_notion_runs([{"plain_text": "x"}])
    assert runs[0]["t"] == "x" and runs[0]["b"] is False
