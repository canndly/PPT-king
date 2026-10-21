from __future__ import annotations

from pathlib import Path

from ppt_king.assets import AssetSink
from ppt_king.ir import ChartBlock, ImageBlock, Slide, TableBlock, TextBlock
from ppt_king.reader import read_pptx


def _read(deck_path: Path, tmp_path: Path):
    sink = AssetSink(root=tmp_path / "out")
    return read_pptx(deck_path, sink), sink


def test_slide_count_and_deck_title(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    assert len(deck.slides) == 5
    assert deck.title == "演示文稿标题"


def test_titles(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    assert deck.slides[0].title == "第一章 绪论"
    assert deck.slides[1].title == "数据表"
    assert deck.slides[4].title == ""


def test_paragraphs_and_bullets(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    body = next(
        block
        for block in deck.slides[0].blocks
        if isinstance(block, TextBlock) and block.role == "body"
    )
    texts = [p.text for p in body.paragraphs]
    assert texts[:3] == ["要点一", "要点二", "普通段落"]
    assert body.paragraphs[1].level == 1
    assert body.paragraphs[2].bullet == "number"
    assert body.paragraphs[0].runs[0].bold is True


def test_notes(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    assert [p.text for p in deck.slides[0].notes] == ["这是备注第一行", "备注第二行"]


def test_table_with_merge(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    table = next(b for b in deck.slides[1].blocks if isinstance(b, TableBlock))
    assert len(table.rows) == 3
    assert table.rows[1][0] == "电压"
    assert (0, 1, 1, 2) in table.merges
    assert (0, 2) in table.spanned


def test_image_saved(deck_path: Path, tmp_path: Path) -> None:
    deck, sink = _read(deck_path, tmp_path)
    images = deck.images
    assert images
    for image in images:
        assert image.asset_id.startswith("IMG")
        assert (sink.root / image.rel_path).is_file()


def test_chart_extracted(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    chart = next(b for b in deck.slides[2].blocks if isinstance(b, ChartBlock))
    assert chart.title == "效率曲线"
    assert chart.categories == ["0.2A", "0.5A", "1.0A"]
    assert chart.series[0][0] == "效率"
    assert chart.series[0][1] == ["80.1", "81.5", "82"]


def test_group_shape_recursed(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    texts = [
        block.paragraphs[0].text
        for block in deck.slides[2].blocks
        if isinstance(block, TextBlock)
    ]
    assert "组合内文字" in texts


def test_empty_slide_has_no_blocks(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    assert deck.slides[4].blocks == []


def test_reading_order_top_then_left(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    positions = [(b.top, b.left) for b in deck.slides[1].blocks]
    assert positions == sorted(positions, key=lambda item: (item[0], item[1]))


def test_slide_images_property(deck_path: Path, tmp_path: Path) -> None:
    deck, _ = _read(deck_path, tmp_path)
    for slide in deck.slides:
        assert slide.images == [b for b in slide.blocks if isinstance(b, ImageBlock)]
        assert len(slide.images) in (0, 1)
    assert isinstance(deck.slides[0], Slide)
