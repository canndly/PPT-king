from __future__ import annotations

from ppt_king.grid import parse_grid
from ppt_king.template_adjust import preview_lines
from ppt_king.templates import catalog_text, load_templates


def test_all_builtin_templates_parse_and_preview() -> None:
    templates = load_templates()

    assert len(templates) >= 9
    for template in templates:
        deck = parse_grid(template.content, source=template.path)
        assert deck.slides, template.name
        lines = preview_lines(deck, 0, 80, 24)
        assert lines, template.name


def test_catalog_contains_all_template_names() -> None:
    templates = load_templates()
    catalog = catalog_text(templates)

    assert "标题页.md" in catalog
    assert "4点2x2文字页.md" in catalog
    assert "公式说明页.md" in catalog
    assert "总结页.md" in catalog
