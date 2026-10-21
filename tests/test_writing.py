from __future__ import annotations

from pathlib import Path

import pytest

from ppt_king.cache import Cache
from ppt_king.config import Settings
from ppt_king.llm import LLMClient
from ppt_king.outline import Outline, OutlinePage, parse_outline
from ppt_king.sources import SourceBundle, SourceImage, scan_sources
from ppt_king.writing import (
    PageResult,
    WritingError,
    assemble_deck,
    build_page_payload,
    clean_output,
    dedupe_duplicate_images,
    dedupe_duplicate_tables,
    select_evidence,
    select_images,
    split_chunks,
    validate_page,
    write_page,
)


def _bundle(tmp_path: Path):
    source = tmp_path / "资料"
    source.mkdir()
    (source / "notes.md").write_text(
        "# 电源设计\n\n输入电压为 24V，输出为 5V。\n\n效率达到 81.2%。\n\n"
        "反激电路适合中小功率场合。\n",
        encoding="utf-8",
    )
    return scan_sources(source, out_dir=tmp_path / "out")


def _settings(tmp_path: Path) -> Settings:
    return Settings(
        api_key="k",
        base_url="http://localhost",
        model="fake",
        concurrency=1,
        cache_dir=tmp_path / "cache",
    )


def test_select_evidence_prefers_relevant_paragraph(tmp_path: Path) -> None:
    bundle = _bundle(tmp_path)
    page = OutlinePage(number=1, title="输入电压", points=["24V 输入"])
    evidence = select_evidence(bundle, page, max_chars=200)
    assert "24V" in evidence
    assert "【资料：notes.md】" in evidence


def test_select_images_filters_unrelated_images(tmp_path: Path) -> None:
    bundle = SourceBundle(
        root=tmp_path,
        out_dir=tmp_path / "out",
        images=[
            SourceImage(
                rel="assets/speed.png",
                path=tmp_path / "speed.png",
                title="speed",
                caption="平地冲刺 3.9 m/s",
            ),
            SourceImage(
                rel="assets/chart.png",
                path=tmp_path / "chart.png",
                title="chart",
                caption="Command Area 实验曲线",
            ),
        ],
    )
    page = OutlinePage(number=1, title="实测运动性能", points=["室内最高速度 3.9 m/s"])
    assert [image.rel for image in select_images(bundle, page)] == ["assets/speed.png"]

    unrelated = OutlinePage(number=2, title="训练流程", points=["PPO 4000 并行 agent"])
    assert select_images(bundle, unrelated) == []


def test_validate_page_requires_coordinates() -> None:
    with pytest.raises(WritingError):
        validate_page(1, "页", "没有坐标的正文", set())


def test_validate_page_rejects_unknown_image() -> None:
    with pytest.raises(WritingError):
        validate_page(1, "页", "@ 1,2 8x4\n![图](assets/evil.png)", set())


def test_validate_page_accepts_known_image() -> None:
    validate_page(1, "页", "@ 1,2 8x4\n![图](assets/a.png)", {"assets/a.png"})


def test_validate_page_with_logo_rejects_title_overlap() -> None:
    with pytest.raises(WritingError, match="第 3 行"):
        validate_page(2, "页面标题", "@ 1,2 16x6\n这是正文", set(), logo=True)


def test_validate_page_with_logo_allows_header_and_body() -> None:
    validate_page(
        2,
        "正文",
        "@ 9,1 8x1 right 20\n页眉\n\n@ 1,3 16x5\n正文",
        set(),
        logo=True,
    )


def test_validate_page_rejects_overlapping_blocks() -> None:
    with pytest.raises(WritingError, match="块重叠"):
        validate_page(
            2,
            "正文",
            "@ 1,3 8x4\n左\n\n@ 5,3 8x4\n右",
            set(),
            logo=True,
        )


def test_write_page(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    def fake_complete(self, system: str, user: str) -> str:
        assert "PPT 内容撰写器" in system
        return "@ 1,2 16x8\n输入电压为 24V。"

    monkeypatch.setattr(LLMClient, "complete", fake_complete)
    settings = _settings(tmp_path)
    outline, bundle = parse_outline("# 演示\n\n## 输入\n\n- 24V\n"), _bundle(tmp_path)

    result = write_page(
        outline.pages[0],
        outline,
        1,
        bundle,
        settings=settings,
        client=LLMClient(settings),
        cache=Cache.create(tmp_path / "cache", enabled=False),
        use_cache=False,
    )

    assert result.ok
    assert result.chunks == ["@ 1,2 16x8\n输入电压为 24V。"]


def test_write_page_rejects_unknown_image_after_retry(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(
        LLMClient,
        "complete",
        lambda self, system, user: "@ 10,3 7x5\n![图](assets/evil.png)",
    )
    settings = _settings(tmp_path)
    outline, bundle = parse_outline("# 演示\n\n## 图\n\n- 架构\n"), _bundle(tmp_path)

    with pytest.raises(WritingError):
        write_page(
            outline.pages[0],
            outline,
            1,
            bundle,
            settings=settings,
            client=LLMClient(settings),
            cache=Cache.create(tmp_path / "cache", enabled=False),
            use_cache=False,
        )


def test_clean_output_strips_outer_fence_but_keeps_code_blocks() -> None:
    assert clean_output("```markdown\n@ 1,2 16x8\n正文\n```") == "@ 1,2 16x8\n正文"
    cleaned = clean_output("@ 1,2 16x8\n```python\nprint(1)\n```")
    assert "```python" in cleaned


def test_page_payload_mentions_logo_layout() -> None:
    page = OutlinePage(number=1, title="页")
    outline = Outline(title="演示", pages=[page])
    with_logo = build_page_payload(page, outline, 1, "证据", [], logo=True)
    without_logo = build_page_payload(page, outline, 1, "证据", [], logo=False)
    assert "正文从第 3 行开始" in with_logo
    assert "正文从第 2 行开始" in without_logo


def test_clean_output_strips_inline_backticks() -> None:
    assert clean_output("公式 `v·ω ≤ μg`，效率 `81.24%`") == "公式 v·ω ≤ μg，效率 81.24%"
    cleaned = clean_output("@ 1,2 16x8\n```python\nprint(1)\n```")
    assert "```python" in cleaned


def test_split_chunks() -> None:
    assert split_chunks("@ 1,2 8x4\n前\n<!-- split -->\n@ 1,2 16x4\n后") == [
        "@ 1,2 8x4\n前",
        "@ 1,2 16x4\n后",
    ]


def test_dedupe_duplicate_tables_keeps_data_page() -> None:
    outline = Outline(
        title="演示",
        pages=[
            OutlinePage(number=1, title="设计要点"),
            OutlinePage(number=2, title="实验数据"),
        ],
    )
    table = "@ 1,6 16x4\n| 指标 | 数值 |\n| --- | --- |\n| 效率 | 81.24 |"
    results = [
        PageResult(
            number=1,
            title="设计要点",
            chunks=[f"@ 1,2 16x3\n- 要点\n\n{table}"],
        ),
        PageResult(number=2, title="实验数据", chunks=[table]),
    ]

    deduped = dedupe_duplicate_tables(outline, results)

    assert "指标" not in deduped[0].chunks[0]
    assert "指标" in deduped[1].chunks[0]


def test_dedupe_duplicate_images_keeps_relevant_page() -> None:
    outline = Outline(
        title="演示",
        pages=[
            OutlinePage(number=1, title="背景"),
            OutlinePage(number=2, title="摩擦约束"),
        ],
    )
    image = "@ 10,3 7x5\n![摩擦约束](assets/friction.png)"
    results = [
        PageResult(number=1, title="背景", chunks=[f"@ 1,2 8x8\n背景\n\n{image}"]),
        PageResult(number=2, title="摩擦约束", chunks=[image]),
    ]

    deduped = dedupe_duplicate_images(outline, results)

    assert "friction.png" not in deduped[0].chunks[0]
    assert "friction.png" in deduped[1].chunks[0]


def test_assemble_deck_uses_placeholders() -> None:
    outline = Outline(
        title="演示",
        pages=[
            OutlinePage(number=1, title="封面", points=["副标题"]),
            OutlinePage(number=2, title="正文", points=[]),
        ],
    )
    results = [
        PageResult(number=1, title="封面", error="失败"),
        PageResult(number=2, title="正文", error="失败"),
    ]
    text = assemble_deck(outline, results)
    assert "# 演示" in text
    assert "center 34" in text
    assert "内容待补充" in text
    assert len(split_chunks(text)) == 1  # 占位内容本身没有 split 标记
