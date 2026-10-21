"""测试夹具：用 python-pptx 现场生成示例 PPTX。"""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


def _add_bullet(paragraph, char: str = "•", numbered: bool = False) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for node in ppr.findall(qn(tag)):
            ppr.remove(node)
    if numbered:
        from lxml import etree

        node = etree.SubElement(ppr, qn("a:buAutoNum"))
        node.set("type", "arabicPeriod")
    else:
        from lxml import etree

        node = etree.SubElement(ppr, qn("a:buChar"))
        node.set("char", char)


def _tiny_png() -> bytes:
    # 1x1 透明 PNG
    import base64

    return base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAwAB/AH"
        "4WfK6AAAAAElFTkSuQmCC"
    )


def build_deck(path: Path) -> Path:
    prs = Presentation()
    prs.core_properties.title = "演示文稿标题"

    # --- 页 1：标题 + 多级列表 + 粗斜体 + 备注 -------------------------
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "第一章 绪论"
    body = slide.placeholders[1].text_frame
    body.text = ""
    p0 = body.paragraphs[0]
    run = p0.add_run()
    run.text = "要点一"
    run.font.bold = True
    _add_bullet(p0)
    p1 = body.add_paragraph()
    p1.text = "要点二"
    p1.level = 1
    _add_bullet(p1)
    p2 = body.add_paragraph()
    p2.text = "普通段落"
    _add_bullet(p2, numbered=True)
    slide.notes_slide.notes_text_frame.text = "这是备注第一行\n备注第二行"

    # --- 页 2：表格（含合并）+ 图片 -----------------------------------
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "数据表"
    rows, cols = 3, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(6), Inches(2)).table
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "数值"
    table.cell(0, 2).text = "单位"
    table.cell(1, 0).text = "电压"
    table.cell(1, 1).text = "5.0"
    table.cell(1, 2).text = "V"
    table.cell(2, 0).text = "电流"
    table.cell(2, 1).text = "1.0"
    table.cell(2, 2).text = "A"
    table.cell(0, 1).merge(table.cell(0, 2))
    slide.shapes.add_picture(
        __import__("io").BytesIO(_tiny_png()), Inches(7), Inches(1.5), Inches(1), Inches(1)
    )

    # --- 页 3：图表 + 组合形状 ----------------------------------------
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "效率曲线"
    data = CategoryChartData()
    data.categories = ["0.2A", "0.5A", "1.0A"]
    data.add_series("效率", (80.1, 81.5, 82.0))
    graphic = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.5), Inches(5), Inches(3), data
    )
    graphic.chart.has_title = True
    graphic.chart.chart_title.text_frame.text = "效率曲线"
    group = slide.shapes.add_group_shape()
    box = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6), Inches(1.5), Inches(2), Inches(1))
    box.text_frame.text = "组合内文字"

    # --- 页 4：纯图片页 -------------------------------------------------
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(
        __import__("io").BytesIO(_tiny_png()), Inches(1), Inches(1), Inches(2), Inches(2)
    )

    # --- 页 5：空页 -----------------------------------------------------
    prs.slides.add_slide(prs.slide_layouts[6])

    prs.save(str(path))
    return path


@pytest.fixture()
def deck_path(tmp_path: Path) -> Path:
    return build_deck(tmp_path / "demo.pptx")
