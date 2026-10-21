"""16×9 定位网格语法：在 Markdown 上用 ``@ 列,行 [宽x高] [对齐]`` 指定块位置。

坐标从 1 开始，``1,1`` 是左上角。每个格子是正方形（762000 EMU = 0.833 英寸），
因为 16:9 画布按 16 列 9 行等分时列宽恰好等于行高。

语法只有三条：

1. ``# 标题``      —— 演示文稿标题（可省略）
2. ``## 页标题``   —— 开始新的一页
3. ``@ 列,行 [宽x高] [left|center|right]`` —— 作用于其后的内容区

**空行不再切块**：一个 ``@`` 到下一个 ``@`` / ``##`` 之间的全部内容属于同一个内容区，
内部按 Markdown 原样保留段落、列表、小标题的层次。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

GRID_COLS = 16
GRID_ROWS = 9
EMU_PER_CELL = 762000  # 12192000 / 16 == 6858000 / 9

ANCHOR_RE = re.compile(
    r"^@\s*(\d{1,2})\s*[,，]\s*(\d{1,2})"
    r"(?:\s*(\d{1,2})\s*[x×*]\s*(\d{1,2}))?"
    r"((?:\s+[^\s]+){0,2})\s*$",
    re.IGNORECASE,
)
SIZE_TOKEN_RE = re.compile(r"^(\d{1,2})$")
IMAGE_RE = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)\s]+)(?:\s+\"[^\"]*\")?\)\s*$")
TABLE_SEP_RE = re.compile(r"^\|[\s:|-]+\|$")
ORDERED_RE = re.compile(r"^(\d+)[.)]\s+(.*)$")
UNORDERED_RE = re.compile(r"^[-*+]\s+(.*)$")
HEADING_RE = re.compile(r"^(#{3,6})\s+(.*)$")
DECK_TITLE_RE = re.compile(r"^#\s+(.*)$")
PAGE_TITLE_RE = re.compile(r"^##(?:\s+(.*))?$")

ALIGN_ALIASES = {"居中": "center", "左": "left", "右": "right"}
CAPTION_PREFIXES = ("图注：", "图注:", "图：", "图:", "注：", "注:")

# 页标题里的自动序号，如 "1. 绪论" / "3、方法"；要求点号后必须有空白，
# 以免误伤 "3.9 m/s" 这类数值。
_TITLE_NUMBER_RE = re.compile(r"^\d{1,2}\s*[.、．)）]\s+(?=\S)")

TEXT = "text"
IMAGE = "image"
TABLE = "table"
QUOTE = "quote"
CODE = "code"

_ECHO_STRIP_RE = re.compile(r"[\s\W_]+", re.UNICODE)
_TITLE_INDEX_RE = re.compile(r"^\d+\s*[.、．)）]\s*")


def _clean_title(title: str) -> str:
    """去掉页标题里手工加的序号（`1. ` / `3、`），渲染时页码由版式负责。"""
    text = title.strip()
    cleaned = _TITLE_NUMBER_RE.sub("", text, count=1)
    return cleaned.strip() or text


def _key(text: str) -> str:
    """用于比较标题与首段是否重复：去掉空白、标点与 Markdown 标记。"""
    return _ECHO_STRIP_RE.sub("", text.replace("*", "").replace("`", "")).lower()


_FOOTER_RUN_RE = re.compile(r"^[^\n]{0,40}[·•|]\s*[^\n]{0,30}$")
_PAGE_NO_RE = re.compile(r"^\d{1,3}$")


def _strip_footer(slide: GridSlide) -> None:
    """去掉页脚跑马灯与页码（导出后常残留在正文末尾）。"""
    for block in reversed(slide.blocks):
        if block.kind != TEXT or not block.flow:
            continue
        removed = 0
        while block.flow and removed < 2:
            tail = block.flow[-1]
            text = tail.text.strip()
            if not text:
                block.flow.pop()
                continue
            if _PAGE_NO_RE.match(text) or _FOOTER_RUN_RE.match(text):
                block.flow.pop()
                removed += 1
                continue
            break
        if not block.flow:
            slide.blocks.remove(block)


def _drop_title_echo(slide: GridSlide) -> None:
    """模型常把页标题再写一遍成首段，这里做确定性去重。"""
    if not slide.title.strip():
        return
    target = _key(_TITLE_INDEX_RE.sub("", slide.title))
    if not target:
        return
    for index, block in enumerate(slide.blocks):
        if block.kind != TEXT or not block.flow:
            continue
        first = block.flow[0]
        if _key(first.text) == target:
            block.flow.pop(0)
            if not block.flow:
                slide.blocks.pop(index)
        return


class GridSyntaxError(ValueError):
    """网格语法不可解析。"""


@dataclass
class Box:
    col: int
    row: int
    colspan: int | None = None
    rowspan: int | None = None
    align: str = "left"
    size: float | None = None  # 明确指定的字号（pt）

    def clamped(self) -> "Box":
        col = min(max(1, self.col), GRID_COLS)
        row = min(max(1, self.row), GRID_ROWS)
        span = self.colspan
        if span is not None:
            span = min(max(1, span), GRID_COLS - col + 1)
        span_y = self.rowspan
        if span_y is not None:
            span_y = min(max(1, span_y), GRID_ROWS - row + 1)
        return Box(col, row, span, span_y, self.align, self.size)

    @property
    def x(self) -> int:
        return (self.col - 1) * EMU_PER_CELL

    @property
    def y(self) -> int:
        return (self.row - 1) * EMU_PER_CELL

    @property
    def width(self) -> int:
        return (self.colspan or (GRID_COLS - self.col + 1)) * EMU_PER_CELL

    @property
    def height(self) -> int:
        return (self.rowspan or (GRID_ROWS - self.row + 1)) * EMU_PER_CELL


@dataclass
class FlowItem:
    """内容区里的一段：普通段落、列表项或小标题。"""

    kind: str  # para | bullet | ordered | heading
    text: str
    level: int = 0


@dataclass
class GridBlock:
    kind: str
    box: Box | None = None
    flow: list[FlowItem] = field(default_factory=list)
    path: str | None = None
    caption: str = ""
    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)

    @property
    def is_positioned(self) -> bool:
        return self.box is not None

    @property
    def preview(self) -> str:
        if self.kind == TABLE:
            return f"表格 {len(self.rows) + 1}×{len(self.header)}"
        if self.kind == IMAGE:
            return f"图片 {self.caption or self.path or ''}"
        if self.kind == CODE:
            return "代码"
        return " ".join(item.text for item in self.flow)[:60]


@dataclass
class GridSlide:
    number: int
    title: str = ""
    blocks: list[GridBlock] = field(default_factory=list)
    notes: str = ""


@dataclass
class GridDeck:
    title: str
    slides: list[GridSlide] = field(default_factory=list)
    source: Path | None = None

    @property
    def positioned_blocks(self) -> int:
        return sum(1 for s in self.slides for b in s.blocks if b.is_positioned)


# --------------------------------------------------------------------------
# 内容区解析
# --------------------------------------------------------------------------
def _split_table_row(line: str) -> list[str]:
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|"):
        body = body[:-1]
    return [cell.strip() for cell in body.split("|")]


def _parse_table(lines: list[str]) -> GridBlock:
    header: list[str] = []
    body_rows: list[list[str]] = []
    for line in lines:
        if TABLE_SEP_RE.match(line.strip()):
            continue
        cells = _split_table_row(line)
        if not header:
            header = cells
        else:
            body_rows.append(cells)
    width = len(header)
    body_rows = [row + [""] * (width - len(row)) for row in body_rows]
    return GridBlock(kind=TABLE, header=header, rows=body_rows)


def _parse_flow(lines: list[str]) -> list[FlowItem]:
    items: list[FlowItem] = []
    paragraph: list[str] = []

    def flush_paragraph() -> None:
        if paragraph:
            text = " ".join(part.strip() for part in paragraph).strip()
            if text:
                items.append(FlowItem("para", text))
            paragraph.clear()

    for line in lines:
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            continue

        heading = HEADING_RE.match(stripped)
        if heading:
            flush_paragraph()
            items.append(FlowItem("heading", heading.group(2).strip()))
            continue

        indent = len(line) - len(line.lstrip(" "))
        level = min(indent // 2, 3)

        ordered = ORDERED_RE.match(stripped)
        if ordered:
            flush_paragraph()
            items.append(FlowItem("ordered", ordered.group(2).strip(), level))
            continue

        unordered = UNORDERED_RE.match(stripped)
        if unordered:
            flush_paragraph()
            items.append(FlowItem("bullet", unordered.group(1).strip(), level))
            continue

        paragraph.append(stripped)

    flush_paragraph()
    return items


def _parse_block(lines: list[str]) -> GridBlock | None:
    while lines and not lines[0].strip():
        lines = lines[1:]
    while lines and not lines[-1].strip():
        lines = lines[:-1]
    if not lines:
        return None

    first = lines[0].strip()

    if first.startswith("```"):
        closed = len(lines) > 1 and lines[-1].strip().startswith("```")
        body = lines[1:-1] if closed else lines[1:]
        return GridBlock(kind=CODE, flow=[FlowItem("para", "\n".join(body))])

    image = IMAGE_RE.match(first)
    if image:
        caption = image.group("alt").strip()
        tail = [line.strip() for line in lines[1:] if line.strip()]
        if tail:
            extra = tail[0]
            if extra.startswith(CAPTION_PREFIXES):
                caption = extra.split("：", 1)[-1].split(":", 1)[-1].strip()
            elif not caption:
                caption = extra
        rel = image.group("path").strip()
        if rel.startswith("./"):
            rel = rel[2:]
        return GridBlock(kind=IMAGE, path=rel, caption=caption)

    if all(line.strip().startswith("|") for line in lines if line.strip()):
        return _parse_table(lines)

    if all(line.strip().startswith(">") for line in lines if line.strip()):
        text = "\n".join(
            line.strip().lstrip(">").strip() for line in lines if line.strip()
        ).strip()
        return GridBlock(kind=QUOTE, flow=[FlowItem("para", text)])

    flow = _parse_flow(lines)
    if not flow:
        return None
    return GridBlock(kind=TEXT, flow=flow)


# --------------------------------------------------------------------------
# 主解析
# --------------------------------------------------------------------------
def parse_grid(text: str, *, source: Path | None = None) -> GridDeck:
    """把网格 Markdown 解析成 GridDeck。"""
    deck_title = source.stem if source else "演示文稿"
    slides: list[GridSlide] = []
    current: GridSlide | None = None
    buffer: list[str] = []
    pending: Box | None = None

    def flush() -> None:
        nonlocal buffer, pending
        if current is not None:
            block = _parse_block(buffer)
            if block is not None:
                block.box = pending
                current.blocks.append(block)
        buffer = []
        pending = None

    def start_page(title: str) -> None:
        nonlocal current
        current = GridSlide(number=len(slides) + 1, title=_clean_title(title))
        slides.append(current)

    for raw in text.replace("\r\n", "\n").replace("\t", "    ").split("\n"):
        line = raw.rstrip()

        page = PAGE_TITLE_RE.match(line)
        if page:
            flush()
            start_page(page.group(1) or "")
            continue

        deck = DECK_TITLE_RE.match(line)
        if deck:
            flush()
            title = deck.group(1).strip()
            if title:
                deck_title = title
            continue

        anchor = ANCHOR_RE.match(line.strip())
        if anchor:
            flush()
            align = "left"
            size: float | None = None
            for token in anchor.group(5).split():
                low = token.lower()
                low = ALIGN_ALIASES.get(low, low)
                if low in ("left", "center", "right"):
                    align = low
                    continue
                numeric = SIZE_TOKEN_RE.match(low)
                if numeric:
                    size = float(numeric.group(1))
            pending = Box(
                col=int(anchor.group(1)),
                row=int(anchor.group(2)),
                colspan=int(anchor.group(3)) if anchor.group(3) else None,
                rowspan=int(anchor.group(4)) if anchor.group(4) else None,
                align=align,
                size=size,
            ).clamped()
            continue

        if current is None:
            if not line.strip():
                continue
            start_page("")

        buffer.append(line)

    flush()
    if not slides:
        raise GridSyntaxError("没有找到任何 `## 页标题`，无法确定页面")
    for slide in slides:
        _drop_title_echo(slide)
        _strip_footer(slide)
    return GridDeck(title=deck_title, slides=slides, source=source)


def load_grid(path: str | Path) -> GridDeck:
    target = Path(path)
    if not target.is_file():
        raise GridSyntaxError(f"找不到文件：{target}")
    return parse_grid(target.read_text(encoding="utf-8-sig"), source=target)


def describe(deck: GridDeck) -> str:
    lines = [f"演示文稿：{deck.title}（{len(deck.slides)} 页，{deck.positioned_blocks} 个定位块）"]
    for slide in deck.slides:
        lines.append(f"  P{slide.number:02d} {slide.title or '(无标题)'}")
        for block in slide.blocks:
            if block.box is not None:
                box = block.box
                span = f" {box.colspan}x{box.rowspan}" if (box.colspan or box.rowspan) else ""
                align = "" if box.align == "left" else f" {box.align}"
                where = f"@{box.col},{box.row}{span}{align}"
            else:
                where = "auto"
            extra = ""
            if block.kind == TEXT:
                counts: dict[str, int] = {}
                for item in block.flow:
                    counts[item.kind] = counts.get(item.kind, 0) + 1
                extra = " " + " ".join(f"{k}×{v}" for k, v in counts.items())
            lines.append(f"     {where:14} {block.kind:6}{extra}  {block.preview[:44]}")
    return "\n".join(lines)
