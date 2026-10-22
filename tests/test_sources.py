from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from ppt_king.sources import SourceError, describe_skipped, discover_outline, scan_sources


def _png(path: Path, color: tuple[int, int, int] = (20, 60, 120)) -> None:
    Image.new("RGB", (64, 36), color).save(path)


def _make_source(root: Path) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    (root / "outline.md").write_text(
        "# 我的汇报\n\n## 封面\n\n- 副标题\n\n## 正文\n\n- 要点\n", encoding="utf-8"
    )
    (root / "notes.md").write_text(
        "# 项目说明\n\n输入电压为 24V。\n\n![架构图](pic.png)\n", encoding="utf-8"
    )
    (root / "extra.txt").write_text("补充资料，效率 81.2%。\n", encoding="utf-8")
    _png(root / "pic.png")
    (root / "unsupported.pdf").write_bytes(b"%PDF-1.4")
    return root


def test_scan_sources_reads_text_and_copies_images(tmp_path: Path) -> None:
    source = _make_source(tmp_path / "资料")
    out = tmp_path / "out"

    bundle = scan_sources(source, out_dir=out)

    assert [doc.rel for doc in bundle.docs] == ["extra.txt", "notes.md"]
    assert bundle.outline == (source / "outline.md").resolve()
    assert bundle.images
    image = bundle.images[0]
    assert image.rel.startswith("assets/src01_")
    assert image.caption == "架构图"
    assert image.path.is_file()
    assert any(item.startswith("unsupported.pdf") for item in bundle.skipped)
    assert "unsupported.pdf" in describe_skipped(bundle)


def test_scan_sources_deduplicates_same_image(tmp_path: Path) -> None:
    source = tmp_path / "资料"
    source.mkdir()
    (source / "a.md").write_text("# A\n\n正文\n", encoding="utf-8")
    _png(source / "one.png")
    (source / "two.png").write_bytes((source / "one.png").read_bytes())

    bundle = scan_sources(source, out_dir=tmp_path / "out")

    assert len(bundle.images) == 1


def test_explicit_outline_wins(tmp_path: Path) -> None:
    source = _make_source(tmp_path / "资料")
    explicit = tmp_path / "custom.md"
    explicit.write_text("# 自定义\n\n## 只有一页\n\n- 要点\n", encoding="utf-8")

    bundle = scan_sources(source, out_dir=tmp_path / "out", outline=explicit)

    assert bundle.outline == explicit.resolve()
    assert all(doc.rel != "outline.md" for doc in bundle.docs)


def test_discover_outline_priority(tmp_path: Path) -> None:
    root = tmp_path / "资料"
    root.mkdir()
    (root / "大纲.md").write_text("# 提纲\n", encoding="utf-8")
    assert discover_outline(root) == (root / "大纲.md").resolve()

    (root / "outline.md").write_text("# 提纲\n", encoding="utf-8")
    assert discover_outline(root) == (root / "outline.md").resolve()


def test_discover_outline_ignores_nested(tmp_path: Path) -> None:
    root = tmp_path / "资料"
    nested = root / "子项目"
    nested.mkdir(parents=True)
    (nested / "OUTLINE.md").write_text("# 子项目大纲\n", encoding="utf-8")

    assert discover_outline(root) is None


def test_missing_outline_raises(tmp_path: Path) -> None:
    source = tmp_path / "资料"
    source.mkdir()
    (source / "a.md").write_text("# A\n正文\n", encoding="utf-8")
    with pytest.raises(SourceError):
        scan_sources(source, out_dir=tmp_path / "out", outline=tmp_path / "nope.md")


def test_no_text_sources_raises(tmp_path: Path) -> None:
    source = tmp_path / "资料"
    source.mkdir()
    _png(source / "only.png")
    with pytest.raises(SourceError):
        scan_sources(source, out_dir=tmp_path / "out")
