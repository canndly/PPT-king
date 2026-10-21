from __future__ import annotations

from pathlib import Path

import pytest

from ppt_king.cache import Cache
from ppt_king.config import Settings
from ppt_king.llm import LLMClient
from ppt_king.outline import (
    OutlineError,
    ensure_cover,
    generate_outline,
    outline_to_markdown,
    parse_outline,
    validate_outline,
)
from ppt_king.sources import scan_sources


def test_parse_outline_with_points() -> None:
    outline = parse_outline(
        "# 演示\n\n## 背景\n\n- 现状\n- 问题\n\n## 方案\n\n1. 第一步\n2. 第二步\n"
    )
    assert outline.title == "演示"
    assert [page.title for page in outline.pages] == ["背景", "方案"]
    assert outline.pages[0].points == ["现状", "问题"]
    assert outline.pages[1].points == ["第一步", "第二步"]


def test_parse_outline_falls_back_to_list() -> None:
    outline = parse_outline("- 第一页\n  - 要点一\n- 第二页\n  - 要点二\n", fallback_title="文件")
    assert outline.title == "文件"
    assert [page.title for page in outline.pages] == ["第一页", "第二页"]
    assert outline.pages[0].points == ["要点一"]


def test_outline_roundtrip() -> None:
    outline = parse_outline("# 标题\n\n## 页一\n\n- 要点\n\n## 页二\n\n- 要点二\n")
    text = outline_to_markdown(outline)
    again = parse_outline(text)
    assert [page.title for page in again.pages] == ["页一", "页二"]
    assert again.pages[0].points == ["要点"]


def test_ensure_cover_prepends_missing_title_page() -> None:
    outline = parse_outline("# 演示\n\n## 背景\n\n- 要点\n")
    covered = ensure_cover(outline, "答辩老师")

    assert [page.title for page in covered.pages] == ["封面", "背景"]
    assert covered.pages[0].points == ["演示", "答辩老师"]
    assert [page.number for page in covered.pages] == [1, 2]


def test_ensure_cover_keeps_existing_cover() -> None:
    outline = parse_outline("# 演示\n\n## 封面\n\n- 副标题\n\n## 背景\n\n- 要点\n")
    assert ensure_cover(outline, "答辩老师") is outline


def test_validate_page_range() -> None:
    outline = parse_outline("## 一\n- a\n")
    assert validate_outline(outline, min_pages=2, max_pages=3)
    assert not validate_outline(outline, min_pages=1, max_pages=3)


def _bundle(tmp_path: Path):
    source = tmp_path / "资料"
    source.mkdir()
    (source / "a.md").write_text("# 资料\n\n关键数据 5V。\n", encoding="utf-8")
    return scan_sources(source, out_dir=tmp_path / "out")


def test_generate_outline(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    calls: list[str] = []

    def fake_complete(self, system: str, user: str) -> str:
        calls.append(system)
        if "资料分析员" in system:
            return "- 关键数据 5V"
        return "# 测试标题\n\n## 封面\n\n- 副标题\n\n## 正文\n\n- 要点\n\n## 总结\n\n- 结论"

    monkeypatch.setattr(LLMClient, "complete", fake_complete)
    settings = Settings(api_key="k", base_url="http://localhost", model="fake", concurrency=1)

    text = generate_outline(
        _bundle(tmp_path),
        settings=settings,
        cache=Cache.create(tmp_path / "cache", enabled=False),
        min_pages=3,
        max_pages=3,
        progress=False,
    )

    assert text.startswith("# 测试标题")
    outline = parse_outline(text)
    assert len(outline.pages) == 3
    assert len(calls) == 2  # 摘要 + 大纲


def test_generate_outline_retries_on_wrong_page_count(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    outputs = ["# 标题\n\n## 只有一页\n\n- 要点", "# 标题\n\n## 一\n\n- a\n\n## 二\n\n- b"]
    calls: list[str] = []

    def fake_complete(self, system: str, user: str) -> str:
        if "资料分析员" in system:
            return "- 摘要"
        calls.append(user)
        return outputs.pop(0)

    monkeypatch.setattr(LLMClient, "complete", fake_complete)
    settings = Settings(api_key="k", base_url="http://localhost", model="fake", concurrency=1)

    text = generate_outline(
        _bundle(tmp_path),
        settings=settings,
        cache=Cache.create(tmp_path / "cache", enabled=False),
        min_pages=2,
        max_pages=2,
        progress=False,
    )

    assert len(parse_outline(text).pages) == 2
    assert len(calls) == 2


def test_generate_outline_rejects_empty(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    def fake_complete(self, system: str, user: str) -> str:
        if "资料分析员" in system:
            return "- 摘要"
        return ""

    monkeypatch.setattr(LLMClient, "complete", fake_complete)
    settings = Settings(api_key="k", base_url="http://localhost", model="fake", concurrency=1)
    with pytest.raises(OutlineError):
        generate_outline(
            _bundle(tmp_path),
            settings=settings,
            cache=Cache.create(tmp_path / "cache", enabled=False),
            min_pages=1,
            max_pages=2,
            progress=False,
        )
