from __future__ import annotations

import json

import pytest

from ppt_king.spec import Block, DeckSpec, SlideSpec, SpecError, extract_json, parse_slide


def test_extract_json_tolerates_fence_and_prose() -> None:
    raw = '这是结果：\n```json\n{"layout": "title", "blocks": []}\n```\n以上。'
    assert extract_json(raw) == {"layout": "title", "blocks": []}


def test_extract_json_rejects_non_json() -> None:
    with pytest.raises(SpecError):
        extract_json("抱歉，我无法完成")


def test_block_from_dict_bullets_accepts_strings() -> None:
    block = Block.from_dict({"type": "bullets", "items": ["甲", "乙"]})
    assert [item.text for item in block.items] == ["甲", "乙"]


def test_block_rejects_unknown_type() -> None:
    with pytest.raises(SpecError):
        Block.from_dict({"type": "video", "text": "x"})


def test_block_rejects_unknown_image_ref() -> None:
    with pytest.raises(SpecError):
        Block.from_dict({"type": "image", "ref": "A9"}, allowed_refs={"A1"})


def test_table_padded_to_consistent_width() -> None:
    block = Block.from_dict(
        {"type": "table", "header": ["a", "b", "c"], "rows": [["1"]]}
    )
    assert block.rows == [["1", "", ""]]


def test_table_requires_header() -> None:
    with pytest.raises(SpecError):
        Block.from_dict({"type": "table", "rows": [["1"]]})


def test_bullets_requires_items() -> None:
    with pytest.raises(SpecError):
        Block.from_dict({"type": "bullets", "items": []})


def test_slide_unknown_layout_falls_back() -> None:
    slide = SlideSpec.from_dict({"layout": "unknown", "title": "t", "blocks": []})
    assert slide.layout == "title_content"


def test_parse_slide_uses_fallback_title() -> None:
    raw = json.dumps({"layout": "two_column", "blocks": [{"type": "paragraph", "text": "x"}]})
    slide = parse_slide(
        raw, number=3, allowed_refs={"A1"}, assets={"A1": "assets/a.png"}, fallback_title="标题"
    )
    assert slide.title == "标题"
    assert slide.number == 3


def test_parse_slide_rejects_foreign_ref() -> None:
    raw = json.dumps({"layout": "two_column", "blocks": [{"type": "image", "ref": "A2"}]})
    with pytest.raises(SpecError):
        parse_slide(raw, number=1, allowed_refs={"A1"}, assets={"A1": "assets/a.png"})


def test_deck_roundtrip(tmp_path) -> None:
    deck = DeckSpec(
        title="演示",
        slides=[
            SlideSpec(
                number=1,
                layout="title",
                title="封面",
                blocks=[Block(type="paragraph", text="副标题")],
                assets={"A1": "assets/a.png"},
                notes="备注",
            )
        ],
    )
    path = deck.save(tmp_path / "spec.json")
    loaded = DeckSpec.load(path)
    assert loaded.to_dict() == deck.to_dict()


def test_parse_slide_drops_duplicate_title() -> None:
    raw = json.dumps(
        {
            "layout": "title_content",
            "blocks": [
                {"type": "paragraph", "text": "一，题目解析"},
                {"type": "paragraph", "text": "正文"},
            ],
        }
    )
    slide = parse_slide(raw, number=2, allowed_refs=set(), assets={}, fallback_title="一，题目解析")
    assert [block.text for block in slide.blocks] == ["正文"]


def test_parse_slide_drops_duplicate_title_bullet() -> None:
    raw = json.dumps(
        {
            "layout": "title_content",
            "blocks": [
                {
                    "type": "bullets",
                    "items": [{"text": "章节标题"}, {"text": "要点"}],
                }
            ],
        }
    )
    slide = parse_slide(raw, number=1, allowed_refs=set(), assets={}, fallback_title="章节标题")
    assert [item.text for item in slide.blocks[0].items] == ["要点"]


def test_parse_slide_drops_duplicate_title_after_image() -> None:
    raw = json.dumps(
        {
            "layout": "two_column",
            "blocks": [
                {"type": "image", "ref": "A1"},
                {"type": "heading", "text": "设计注意事项"},
                {"type": "paragraph", "text": "正文"},
            ],
        }
    )
    slide = parse_slide(
        raw,
        number=6,
        allowed_refs={"A1"},
        assets={"A1": "assets/a.png"},
        fallback_title="设计注意事项",
    )
    assert [block.type for block in slide.blocks] == ["image", "paragraph"]


def test_deck_load_rejects_empty_slides(tmp_path) -> None:
    target = tmp_path / "bad.json"
    target.write_text('{"slides": []}', encoding="utf-8")
    with pytest.raises(SpecError):
        DeckSpec.load(target)
