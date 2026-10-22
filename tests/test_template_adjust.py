from __future__ import annotations

from pathlib import Path

from ppt_king.cli import _build_parser
from ppt_king.grid import parse_grid
from ppt_king.template_adjust import (
    _display_width,
    choose_template,
    discover_templates,
    fit_preview_size,
    preview_lines,
    terminal_hint,
)


def test_discover_templates_only_markdown(tmp_path: Path) -> None:
    nested = tmp_path / "nested"
    nested.mkdir()
    (tmp_path / "b.md").write_text("## B\n", encoding="utf-8")
    (tmp_path / "a.markdown").write_text("## A\n", encoding="utf-8")
    (tmp_path / "ignore.txt").write_text("x", encoding="utf-8")
    (nested / "c.md").write_text("## C\n", encoding="utf-8")

    found = discover_templates(tmp_path)

    assert [path.name for path in found] == ["a.markdown", "b.md", "c.md"]


def test_choose_template_by_number(tmp_path: Path) -> None:
    first = tmp_path / "a.md"
    second = tmp_path / "b.md"
    output: list[str] = []

    selected = choose_template(
        [first, second],
        input_fn=lambda prompt: "2",
        output_fn=output.append,
    )

    assert selected == second
    assert any("a.md" in line for line in output)


def test_preview_lines_contains_title_text_and_image() -> None:
    deck = parse_grid(
        "## 标题页\n\n"
        "@ 2,4-15,4 center 34\n**主标题**\n\n"
        "@ 2,6-15,6 center 18\n副标题\n\n"
        "@ 2,8-7,8\n![示意图](assets/a.png)\n"
    )

    lines = preview_lines(deck, 0, 80, 24)
    text = "\n".join(lines)

    assert "标题页" in text
    assert "主标题" in text
    assert "副标题" in text
    assert "[图片]" in text
    assert all(_display_width(line) <= 80 for line in lines)


def test_preview_lines_keeps_compact_single_row_blocks() -> None:
    deck = parse_grid(
        "## 页\n\n@ 2,3-9,3 54\n01\n\n@ 2,5-9,5 32\n章节标题"
    )

    text = "\n".join(preview_lines(deck, 0, 80, 24))

    assert "01" in text
    assert "章节标题" in text


def test_preview_lines_respects_center_alignment() -> None:
    deck = parse_grid("## 页\n\n@ 2,4-15,4 center 34\n居中标题")

    lines = preview_lines(deck, 0, 80, 24)
    line = next(line for line in lines if "居中标题" in line)
    left = line.index("居中标题")

    assert left > 5
    assert 80 - (left + _display_width("居中标题")) > 5


def test_terminal_hint_checks_size() -> None:
    assert terminal_hint(100, 30) == ""
    assert "80×24" in terminal_hint(60, 30)
    assert "80×24" in terminal_hint(80, 20)


def test_fit_preview_size_keeps_16_9_in_tall_terminal() -> None:
    width, height, offset_x, offset_y = fit_preview_size(100, 50)

    assert height < 48
    assert offset_y > 0
    assert abs(width / (height * 2) - 16 / 9) < 0.03
    assert offset_x >= 0


def test_fit_preview_size_keeps_16_9_in_wide_terminal() -> None:
    width, height, offset_x, offset_y = fit_preview_size(140, 24)

    assert width < 139
    assert offset_x > 0
    assert abs(width / (height * 2) - 16 / 9) < 0.03
    assert offset_y >= 0


def test_template_adjust_flag_is_global() -> None:
    args = _build_parser().parse_args(["--template-adjust"])

    assert args.template_adjust is True
    assert args.command is None
