def rt(text, **ann):
    """Build a minimal Notion rich_text entry."""
    a = {"bold": False, "italic": False, "underline": False,
         "strikethrough": False, "code": False}
    a.update(ann)
    return {"plain_text": text, "annotations": a, "href": ann.get("href")}


def block(btype, rich=None, **extra):
    payload = {"rich_text": rich or []}
    payload.update(extra)
    return {"type": btype, btype: payload, "has_children": False}


def convert(blocks):
    from word_generator import _notion_blocks_to_content
    return _notion_blocks_to_content(blocks)


def test_heading_relevel():
    out, _ = convert([
        block("heading_1", [rt("One")]),
        block("heading_2", [rt("Two")]),
        block("heading_3", [rt("Three")]),
    ])
    assert [b["type"] for b in out] == ["heading2", "heading3", "heading4"]


def test_numbered_list_preserved():
    out, _ = convert([
        block("numbered_list_item", [rt("first")]),
        block("numbered_list_item", [rt("second")]),
    ])
    assert [b["type"] for b in out] == ["numbered", "numbered"]
    assert out[0]["text"] == "first"


def test_todo_block():
    out, _ = convert([block("to_do", [rt("task")], checked=True)])
    assert out[0]["type"] == "todo" and out[0]["checked"] is True


def test_paragraph_carries_runs():
    out, _ = convert([{"type": "paragraph", "paragraph": {"rich_text": [
        rt("normal "), rt("bold", bold=True)]}, "has_children": False}])
    assert out[0]["type"] == "paragraph"
    assert out[0]["runs"][1] == {"t": "bold", "b": True, "i": False, "u": False,
                                 "s": False, "c": False, "link": None}


def test_code_block():
    out, _ = convert([block("code", [rt("x = 1")], language="python", caption=[])])
    assert out[0] == {"type": "code_block", "text": "x = 1"}


def test_image_block():
    out, _ = convert([{"type": "image", "image": {
        "type": "file", "file": {"url": "https://x/img.png"},
        "caption": [rt("a chart")]}, "has_children": False}])
    assert out[0] == {"type": "image", "url": "https://x/img.png", "caption": "a chart"}


def test_bookmark_becomes_link_paragraph():
    out, _ = convert([{"type": "bookmark", "bookmark": {
        "url": "https://evye.co", "caption": []}, "has_children": False}])
    assert out[0]["type"] == "paragraph"
    assert out[0]["runs"][0]["link"] == "https://evye.co"


def test_unknown_block_never_silently_dropped():
    out, _ = convert([{"type": "some_future_type", "some_future_type": {},
                       "has_children": False}])
    assert len(out) == 1
    assert "[Unsupported block: some_future_type]" in out[0]["text"]


def test_child_page_placeholder():
    out, _ = convert([{"type": "child_page", "child_page": {"title": "Sub Doc"},
                       "has_children": False}])
    assert "[Sub-page: Sub Doc]" in out[0]["text"]


def test_nested_bullets_and_todo_children():
    inner = block("bulleted_list_item", [rt("child")])
    outer = block("bulleted_list_item", [rt("parent")])
    outer["children"] = [inner]
    out, _ = convert([outer])
    assert out[0] == {"type": "bullet", "text": "parent", "level": 1,
                      "runs": out[0]["runs"]}
    assert out[1]["level"] == 2


def test_wide_table_uses_generic_table():
    tbl = {"type": "table", "table": {"has_column_header": True},
           "has_children": True,
           "_rows": [
               {"type": "table_row", "table_row": {"cells": [
                   [rt("A")], [rt("B")], [rt("C")], [rt("D")]]}},
               {"type": "table_row", "table_row": {"cells": [
                   [rt("1")], [rt("2")], [rt("3")], [rt("4")]]}},
           ]}
    out, _ = convert([tbl])
    assert out[0]["type"] == "generic_table"
    assert out[0]["headers"] == ["A", "B", "C", "D"]
    assert out[0]["rows"] == [["1", "2", "3", "4"]]


def test_one_column_table_not_dropped():
    tbl = {"type": "table", "table": {"has_column_header": False},
           "has_children": True,
           "_rows": [{"type": "table_row", "table_row": {"cells": [[rt("only")]]}}]}
    out, _ = convert([tbl])
    assert out[0]["type"] == "generic_table"
    assert out[0]["rows"] == [["only"]]
