from __future__ import annotations

from ppt_king import markdown as md
from ppt_king.ir import EmbedBlock, ImageBlock, Paragraph, Run, Slide


def _image(asset_id: str = "IMG1", rel: str = "assets/a.png", alt: str = "图") -> ImageBlock:
    return ImageBlock(asset_id=asset_id, rel_path=rel, alt=alt)


def test_normalize_strips_fence_and_trailing_rule() -> None:
    raw = "```markdown\n## 1. 标题\n正文\n```\n\n---"
    assert md.normalize(raw) == "## 1. 标题\n\n正文"


def test_normalize_inserts_blank_line_after_heading() -> None:
    assert md.normalize("## 1. 标题\n正文") == "## 1. 标题\n\n正文"


def test_validate_ok() -> None:
    assert md.validate_page("## 1. 标题\n\n{{IMG1}}", {"IMG1"}) == []


def test_validate_rejects_raw_path_and_unknown_placeholder() -> None:
    problems = md.validate_page("## 1. 标题\n![](assets/a.png)\n{{IMG9}}", {"IMG1"})
    assert any("路径" in p for p in problems)
    assert any("IMG9" in p for p in problems)


def test_validate_rejects_fence_and_missing_heading() -> None:
    problems = md.validate_page("```\n正文\n```", set())
    assert any("围栏" in p for p in problems)
    assert any("标题行" in p for p in problems)


def test_resolve_placeholders() -> None:
    page = "前\n{{IMG1}}\n{{FILE1}}\n后"
    resolved = md.resolve_placeholders(
        page,
        {"IMG1": _image(alt="电路图")},
        {"FILE1": EmbedBlock(asset_id="FILE1", rel_path="media/a.mp4", label="视频")},
    )
    assert "![电路图](assets/a.png)" in resolved
    assert "[视频](media/a.mp4)" in resolved


def test_missing_assets_detected_and_appended() -> None:
    slide = Slide(number=1, blocks=[_image("IMG1"), _image("IMG2", "assets/b.png")])
    page = "## 1. 标题\n{{IMG1}}"
    assert md.missing_asset_ids(page, slide) == ["IMG2"]
    assert md.append_missing(page, ["IMG2"]).strip().endswith("{{IMG2}}")


def test_ensure_heading_when_model_omits_it() -> None:
    assert md.ensure_heading("正文", 3, "标题") == "## 3. 标题\n\n正文"
    assert md.ensure_heading("正文", 3, "") == "## 幻灯片 3\n\n正文"


def test_assemble_joins_pages() -> None:
    document = md.assemble("总标题", ["## 1. A", "## 2. B"])
    assert document.startswith("# 总标题\n\n")
    assert "\n\n---\n\n" in document


def test_paragraph_inline_markers() -> None:
    from ppt_king.serialize import slide_payload

    slide = Slide(
        number=2,
        title="标题",
        blocks=[
            __import__("ppt_king.ir", fromlist=["TextBlock"]).TextBlock(
                paragraphs=[
                    Paragraph(runs=[Run("粗", bold=True), Run("斜", italic=True), Run("行内换行\n")]),
                    Paragraph(runs=[Run("第二段")], level=1, bullet="char"),
                ],
                role="body",
            )
        ],
    )
    payload = slide_payload(slide, 4)
    assert "<b>粗</b>" in payload
    assert "<i>斜</i>" in payload
    assert "lvl=1,bullet=char" in payload
    assert "幻灯片 2/4" in payload
