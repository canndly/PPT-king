"""把 IR 序列化成紧凑文本载荷，供大模型转换成 Markdown。"""

from __future__ import annotations

import re

from .ir import (
    ChartBlock,
    EmbedBlock,
    ImageBlock,
    Paragraph,
    Slide,
    TableBlock,
    TextBlock,
)

_TRAILING_BR = re.compile(r"(?:\s*<br>\s*)+$")

ROLE_LABELS = {
    "title": "标题",
    "subtitle": "副标题",
    "body": "正文",
    "textbox": "文本框",
}


def _inline(paragraph: Paragraph) -> str:
    parts: list[str] = []
    for run in paragraph.runs:
        text = run.text.replace("\n", "<br>")
        if run.code:
            text = f"<code>{text}</code>"
        if run.bold and run.italic:
            text = f"<b><i>{text}</i></b>"
        elif run.bold:
            text = f"<b>{text}</b>"
        elif run.italic:
            text = f"<i>{text}</i>"
        parts.append(text)
    return _TRAILING_BR.sub("", "".join(parts)).strip()


def _attrs(paragraph: Paragraph) -> str:
    attrs = [f"lvl={paragraph.level}"]
    if paragraph.bullet:
        attrs.append(f"bullet={paragraph.bullet}")
    return ",".join(attrs)


def _text_block_lines(block: TextBlock, index: int) -> list[str]:
    role = ROLE_LABELS.get(block.role, block.role)
    header = f"[{index}] {role}"
    if block.placeholder:
        header += f" (占位符={block.placeholder})"
    lines = [header]
    for paragraph in block.paragraphs:
        body = _inline(paragraph)
        if not body.strip():
            continue
        lines.append(f"  <p {_attrs(paragraph)}> {body}")
    return lines


def _table_lines(block: TableBlock, index: int) -> list[str]:
    rows = len(block.rows)
    cols = max((len(row) for row in block.rows), default=0)
    lines = [f"[{index}] 表格 {rows}x{cols}"]
    for row_index, col_index, row_span, col_span in block.merges:
        lines.append(f"  合并: 行{row_index}列{col_index} 跨{row_span}行{col_span}列")
    for row_index, row in enumerate(block.rows):
        cells = [cell.replace("\n", "<br>") for cell in row]
        lines.append(f"  行{row_index}: " + " | ".join(cells))
    return lines


def _image_lines(block: ImageBlock, index: int) -> list[str]:
    alt = block.alt.replace("\n", " ")
    return [f"[{index}] 图片 {block.asset_id} alt={alt} 尺寸={block.width}x{block.height}"]


def _chart_lines(block: ChartBlock, index: int) -> list[str]:
    lines = [f"[{index}] 图表 {block.chart_type}"]
    if block.title:
        lines.append(f"  标题: {block.title}")
    if block.categories:
        lines.append("  分类: " + " | ".join(block.categories))
    for name, values in block.series:
        lines.append(f"  系列 {name}: " + " | ".join(values))
    if not block.categories and not block.series:
        lines.append("  (无可用数据)")
    return lines


def _embed_lines(block: EmbedBlock, index: int) -> list[str]:
    return [f"[{index}] 附件 {block.asset_id} 类型={block.kind} 名称={block.label} 相对路径={block.rel_path}"]


def slide_payload(slide: Slide, total: int, include_notes: bool = True) -> str:
    """生成单页载荷文本。"""
    lines = [f"幻灯片 {slide.number}/{total}"]
    if slide.title:
        lines.append(f"页标题(供参考): {slide.title}")

    index = 0
    for block in slide.blocks:
        index += 1
        if isinstance(block, TextBlock):
            lines.extend(_text_block_lines(block, index))
        elif isinstance(block, TableBlock):
            lines.extend(_table_lines(block, index))
        elif isinstance(block, ImageBlock):
            lines.extend(_image_lines(block, index))
        elif isinstance(block, ChartBlock):
            lines.extend(_chart_lines(block, index))
        elif isinstance(block, EmbedBlock):
            lines.extend(_embed_lines(block, index))

    if include_notes and slide.notes:
        lines.append("[备注]")
        for paragraph in slide.notes:
            body = _inline(paragraph).strip()
            if body:
                lines.append(f"  {body}")

    if len(lines) <= 1:
        lines.append("(本页无文字内容)")
    return "\n".join(lines)
