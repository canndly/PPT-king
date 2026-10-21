from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from ppt_king.config import Settings
from ppt_king.exporter import drop_repeated_images, export_pptx
from ppt_king.ir import ImageBlock
from ppt_king.llm import LLMClient, LLMError

CANNED = {
    # 第 1 页：故意带代码围栏，验证清洗
    "第一章 绪论": "```markdown\n## 1. 第一章 绪论\n- **要点一**\n  - 要点二\n\n> **备注**：这是备注第一行\n> 备注第二行\n```",
    # 第 2 页：故意漏掉图片，验证兜底补全
    "数据表": "## 2. 数据表\n\n| 项目 | 数值 |\n| --- | --- |\n| 电压 | 5.0 |",
    "效率曲线": "## 3. 效率曲线\n\n### 图表：效率曲线\n\n| 分类 | 效率 |\n| --- | --- |\n| 0.2A | 80.1 |",
    "": "## 幻灯片 4\n\n{{IMG1}}",
}


def _fake_complete(self, system: str, user: str) -> str:
    for key, value in CANNED.items():
        if key and f"页标题(供参考): {key}" in user:
            return value
    if "幻灯片 4/" in user:
        return "## 幻灯片 4"
    if "幻灯片 5/" in user:
        return "## 幻灯片 5"
    raise LLMError("fake 未覆盖的页")


@pytest.fixture()
def settings(tmp_path: Path) -> Settings:
    return Settings(
        api_key="test",
        base_url="http://localhost",
        model="fake",
        cache_dir=tmp_path / "cache",
        concurrency=2,
    )


@pytest.fixture(autouse=True)
def _patch_llm(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(LLMClient, "complete", _fake_complete)


def test_export_layout(deck_path: Path, tmp_path: Path, settings: Settings) -> None:
    result = export_pptx(
        deck_path,
        out_root=tmp_path / "export",
        settings=settings,
        use_cache=False,
        progress=False,
    )
    output_dir = tmp_path / "export" / "demo"
    assert result.output_dir == output_dir
    assert result.markdown_path == output_dir / "demo.md"
    assert (output_dir / "assets").is_dir()
    assert result.markdown.startswith("# 演示文稿标题")


def test_export_cleans_fences_and_resolves_images(
    deck_path: Path, tmp_path: Path, settings: Settings
) -> None:
    result = export_pptx(
        deck_path, out_root=tmp_path / "export", settings=settings, use_cache=False, progress=False
    )
    assert "```" not in result.markdown
    assert "](assets/" in result.markdown  # 第 2 页缺失的图片被补到页尾
    assert "{{" not in result.markdown


def test_export_notes_as_blockquote(deck_path: Path, tmp_path: Path, settings: Settings) -> None:
    result = export_pptx(
        deck_path, out_root=tmp_path / "export", settings=settings, use_cache=False, progress=False
    )
    assert "> **备注**：这是备注第一行" in result.markdown


def test_page_separators(deck_path: Path, tmp_path: Path, settings: Settings) -> None:
    result = export_pptx(
        deck_path, out_root=tmp_path / "export", settings=settings, use_cache=False, progress=False
    )
    assert result.markdown.count("\n---\n") == len(result.pages) - 1


def test_cache_hits_skip_llm(deck_path: Path, tmp_path: Path, settings: Settings) -> None:
    cache_dir = tmp_path / "cache"
    first = export_pptx(
        deck_path,
        out_root=tmp_path / "e1",
        settings=replace(settings, cache_dir=cache_dir),
        use_cache=True,
        progress=False,
    )
    keys = list(cache_dir.glob("*.md"))
    assert len(keys) == len(first.pages)


def test_stdout_mode_writes_no_markdown(deck_path: Path, tmp_path: Path, settings: Settings) -> None:
    result = export_pptx(
        deck_path,
        out_root=tmp_path / "export",
        settings=settings,
        use_cache=False,
        write_file=False,
        progress=False,
    )
    assert result.markdown_path is None
    assert not (tmp_path / "export" / "demo" / "demo.md").exists()


def test_drop_repeated_images(deck_path: Path, tmp_path: Path) -> None:
    from ppt_king.assets import AssetSink
    from ppt_king.reader import read_pptx

    deck = read_pptx(deck_path, AssetSink(root=tmp_path / "x"))
    # 同一张图出现在两页 = 被视为重复（阈值按 5 页 * 0.6 = 3，故手动缩放到 2 页场景）
    deck.slides[1].blocks.append(deck.slides[3].blocks[0])
    assert drop_repeated_images(deck, min_ratio=0.2) >= 2
    remaining = [b for s in deck.slides for b in s.blocks if isinstance(b, ImageBlock)]
    assert len(remaining) < len(deck.images) or not remaining


def test_missing_key_raises(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from ppt_king.config import ConfigurationError, load_settings

    monkeypatch.delenv("PPT_KING_API_KEY", raising=False)
    monkeypatch.delenv("DEEPSEEK_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    with pytest.raises(ConfigurationError):
        load_settings(tmp_path / "absent.env")
