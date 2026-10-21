from __future__ import annotations

import json
from pathlib import Path

import pytest

from ppt_king.cache import Cache
from ppt_king.config import Settings
from ppt_king.grid import parse_grid
from ppt_king.grid_ai import (
    SourcePage,
    _clean,
    _split_output,
    layout_markdown,
    plan_page,
    split_pages,
)
from ppt_king.llm import LLMClient

SOURCE = """# 演示标题

## 第一页

正文一。

![图](assets/a.png)

## 第二页

正文二。
"""


def test_split_pages() -> None:
    plan = split_pages(SOURCE)
    assert plan.title == "演示标题"
    assert [page.title for page in plan.pages] == ["第一页", "第二页"]
    assert [page.number for page in plan.pages] == [1, 2]
    assert "正文一" in plan.pages[0].body
    assert "assets/a.png" in plan.pages[0].body


def test_split_pages_without_deck_title() -> None:
    plan = split_pages("## 只有一页\n\n正文\n", fallback_title="文件名")
    assert plan.title == "文件名"
    assert len(plan.pages) == 1


def test_clean_strips_fence_and_echo_heading() -> None:
    assert _clean("```markdown\n@ 1,2 16x8\n正文\n```") == "@ 1,2 16x8\n正文"
    assert _clean("## 第一页\n@ 1,2 16x8\n正文") == "@ 1,2 16x8\n正文"


def test_split_output_uses_marker() -> None:
    parts = _split_output("@ 1,2 8x4\n前半\n<!-- split -->\n@ 1,2 16x8\n后半")
    assert len(parts) == 2
    assert parts[0].startswith("@ 1,2 8x4")


def test_layout_markdown_per_page(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    calls: list[str] = []

    def fake_complete(self, system: str, user: str) -> str:
        calls.append(user)
        if "第一页" in user:
            return "@ 1,2 8x6\n正文一。\n\n@ 10,3 7x5\n![图](assets/a.png)"
        return "@ 1,2 16x8\n正文二。"

    monkeypatch.setattr(LLMClient, "complete", fake_complete)

    source = tmp_path / "deck.md"
    source.write_text(SOURCE, encoding="utf-8")
    settings = Settings(api_key="k", base_url="http://localhost", model="fake", cache_dir=tmp_path / "c")

    text = layout_markdown(source, settings=settings, use_cache=False, progress=False)

    assert len(calls) == 2  # 逐页调用
    assert "# 演示标题" in text
    assert "## 第一页" in text and "## 第二页" in text
    assert "@ 10,3 7x5" in text

    deck = parse_grid(text)
    assert len(deck.slides) == 2
    assert all(block.is_positioned for slide in deck.slides for block in slide.blocks)


def test_layout_retries_on_bad_output(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    outputs = ["这不是网格", "@ 1,2 16x8\n正文"]

    def fake_complete(self, system: str, user: str) -> str:
        return outputs.pop(0)

    monkeypatch.setattr(LLMClient, "complete", fake_complete)

    settings = Settings(api_key="k", base_url="http://localhost", model="fake")
    page = SourcePage(number=1, title="页", body="正文")
    chunks = plan_page(
        LLMClient(settings), Cache.create(tmp_path / "c", enabled=False), settings, page, 1
    )
    assert chunks == ["@ 1,2 16x8\n正文"]


def test_split_marker_creates_continuation_page(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    def fake_complete(self, system: str, user: str) -> str:
        return "@ 1,2 16x4\n前半\n<!-- split -->\n@ 1,2 16x4\n后半"

    monkeypatch.setattr(LLMClient, "complete", fake_complete)

    source = tmp_path / "d.md"
    source.write_text("## 长页\n\n内容\n", encoding="utf-8")
    settings = Settings(api_key="k", base_url="http://localhost", model="fake")
    text = layout_markdown(source, settings=settings, use_cache=False, progress=False)

    assert "## 长页" in text
    assert "## 长页（续）" in text
    assert len(parse_grid(text).slides) == 2
