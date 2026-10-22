from __future__ import annotations

from pathlib import Path

import pytest

from ppt_king.cli import _build_parser
from ppt_king.grid import parse_grid
from ppt_king.template_import import extract_templates, write_templates

TEMPLATE_PPTX = Path("static/template.pptx")
pytestmark = pytest.mark.skipif(
    not TEMPLATE_PPTX.is_file(), reason="需要 static/template.pptx 才能做反向导入"
)


def test_extract_templates_from_builtin_pptx() -> None:
    results = extract_templates(TEMPLATE_PPTX)

    assert len(results) >= 9
    assert "章节页" in {item.title for item in results}
    for item in results:
        deck = parse_grid(item.markdown)
        assert deck.slides, item.title


def test_write_templates_exports_markdown_and_images(tmp_path: Path) -> None:
    results = write_templates(TEMPLATE_PPTX, tmp_path, force=True)

    assert len(results) >= 9
    assert (tmp_path / "章节页.md").is_file()
    assert (tmp_path / "assets" / "章节页_img01.png").is_file()
    chapter = (tmp_path / "章节页.md").read_text(encoding="utf-8")
    assert "assets/章节页_img01.png" in chapter
    assert "@2,3-7,3" in chapter


def test_template_import_cli_defaults() -> None:
    args = _build_parser().parse_args(["template-import", "static/template.pptx"])

    assert args.command == "template-import"
    assert args.out == "out/templates-import"
    assert args.force is False
