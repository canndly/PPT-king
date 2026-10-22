"""模板编写模式：选择 static/templates 下的模板，保存后实时预览。"""

from __future__ import annotations

import shutil
import sys
import time
import unicodedata
from pathlib import Path

from .grid import (
    CODE,
    FORMULA,
    GRID_COLS,
    GRID_ROWS,
    IMAGE,
    QUOTE,
    TABLE,
    TEXT,
    GridBlock,
    GridDeck,
    GridSyntaxError,
    load_grid,
)
from .grid_render import _normalize_formula, _segments
from .templates import DEFAULT_TEMPLATE_DIR, discover_templates


def choose_template(
    paths: list[Path],
    *,
    input_fn=input,
    output_fn=print,
) -> Path:
    """在终端列出模板并让用户选择。"""
    if not paths:
        raise FileNotFoundError("模板目录中没有 Markdown 文件")
    output_fn("可编辑模板：")
    for index, path in enumerate(paths, 1):
        output_fn(f"  {index}. {path.name}")
    while True:
        raw = input_fn(f"选择模板 [1-{len(paths)}]（回车=1）：").strip()
        if not raw:
            return paths[0]
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(paths):
                return paths[choice - 1]
        output_fn("输入无效，请重新输入序号。")


# ---------------------------------------------------------------------------
# 终端预览
# ---------------------------------------------------------------------------
PREVIEW_RATIO = 16 / 9
PREVIEW_MIN_WIDTH = 40
PREVIEW_MIN_HEIGHT = 12


def terminal_hint(cols: int, rows: int) -> str:
    """终端尺寸不足时给出可读提示。"""
    if cols >= 80 and rows >= 24:
        return ""
    return "建议终端至少 80×24（推荐 96×27），预览固定 16:9"


def fit_preview_size(cols: int, rows: int) -> tuple[int, int, int, int]:
    """在终端可用区内计算 16:9 预览画布，返回宽、高、左偏移、上偏移。"""
    available_w = max(PREVIEW_MIN_WIDTH, cols - 1)
    available_h = max(PREVIEW_MIN_HEIGHT, rows - 2)
    if available_w / (available_h * 2) > PREVIEW_RATIO:
        width = max(
            PREVIEW_MIN_WIDTH,
            min(available_w, round(available_h * 2 * PREVIEW_RATIO)),
        )
        height = available_h
    else:
        height = max(
            PREVIEW_MIN_HEIGHT,
            min(available_h, round(available_w / (2 * PREVIEW_RATIO))),
        )
        width = available_w
    offset_x = max(0, (cols - 1 - width) // 2)
    offset_y = max(0, (rows - 2 - height) // 2)
    return width, height, offset_x, offset_y


def _char_width(char: str) -> int:
    if unicodedata.combining(char):
        return 0
    return 2 if unicodedata.east_asian_width(char) in "WF" else 1


def _display_width(text: str) -> int:
    return sum(_char_width(char) for char in text)


def _truncate(text: str, width: int) -> str:
    if width <= 0:
        return ""
    out: list[str] = []
    used = 0
    for char in text:
        size = _char_width(char)
        if used + size > width:
            break
        out.append(char)
        used += size
    return "".join(out)


def _wrap(text: str, width: int) -> list[str]:
    if width <= 0:
        return [""]
    lines: list[str] = []
    current: list[str] = []
    used = 0
    for char in text:
        size = _char_width(char)
        if size == 0:
            current.append(char)
            continue
        if used + size > width and current:
            lines.append("".join(current))
            current = []
            used = 0
        current.append(char)
        used += size
    if current or not lines:
        lines.append("".join(current))
    return lines


class _Canvas:
    def __init__(self, width: int, height: int) -> None:
        self.width = width
        self.height = height
        self.rows: list[list[str]] = [[" "] * width for _ in range(height)]

    def put(self, x: int, y: int, text: str, *, limit: int | None = None) -> None:
        if not (0 <= y < self.height):
            return
        max_x = self.width if limit is None else min(self.width, limit)
        cursor = max(0, int(x))
        for char in text:
            size = _char_width(char)
            if size == 0:
                continue
            if cursor + size > max_x:
                break
            self.rows[y][cursor] = char
            if size == 2 and cursor + 1 < self.width:
                self.rows[y][cursor + 1] = ""
            cursor += size

    def rect(self, x1: int, y1: int, x2: int, y2: int) -> None:
        if x2 - x1 < 1 or y2 - y1 < 1:
            return
        self.put(x1, y1, "┌")
        self.put(x2, y1, "┐")
        self.put(x1, y2, "└")
        self.put(x2, y2, "┘")
        for x in range(x1 + 1, x2):
            self.put(x, y1, "─")
            self.put(x, y2, "─")
        for y in range(y1 + 1, y2):
            self.put(x1, y, "│")
            self.put(x2, y, "│")

    def lines(self) -> list[str]:
        rendered: list[str] = []
        for row in self.rows:
            line = "".join(char for char in row if char != "")
            padding = max(0, self.width - _display_width(line))
            rendered.append(line + " " * padding)
        return rendered


def _plain(text: str) -> str:
    return "".join(segment.text for segment in _segments(text))


def _block_lines(block: GridBlock) -> list[str]:
    if block.kind == TEXT:
        result: list[str] = []
        for item in block.flow:
            text = _plain(item.text)
            if item.kind == "bullet":
                text = "• " + text
            elif item.kind == "ordered":
                text = "· " + text
            result.append(text)
        return result or [block.preview]
    if block.kind == FORMULA:
        text = block.flow[0].text if block.flow else block.preview
        return [f"ƒ {_normalize_formula(text)}"]
    if block.kind == IMAGE:
        caption = block.caption or Path(block.path or "").name
        return [f"[图片] {caption}"]
    if block.kind == TABLE:
        header = " | ".join(block.header)
        return [f"[表格 {len(block.rows)}×{len(block.header)}]", header]
    if block.kind == QUOTE:
        text = block.flow[0].text if block.flow else block.preview
        return [f"“{text}”"]
    if block.kind == CODE:
        text = block.flow[0].text if block.flow else ""
        return ["[代码]", *text.splitlines()[:3]]
    return [block.preview]


def _map_box(box, width: int, height: int) -> tuple[int, int, int, int]:
    inner_w = width - 2
    inner_h = height - 2

    def x_of(col: int) -> int:
        return 1 + round((col - 1) / GRID_COLS * inner_w)

    def y_of(row: int) -> int:
        return 1 + round((row - 1) / GRID_ROWS * inner_h)

    span_x = box.colspan or (GRID_COLS - box.col + 1)
    span_y = box.rowspan or (GRID_ROWS - box.row + 1)
    x1 = x_of(box.col)
    x2 = x_of(box.col + span_x) - 1
    y1 = y_of(box.row)
    y2 = y_of(box.row + span_y) - 1
    return (
        max(1, min(width - 2, x1)),
        max(1, min(height - 2, y1)),
        max(1, min(width - 2, x2)),
        max(1, min(height - 2, y2)),
    )


def preview_lines(
    deck: GridDeck, index: int, width: int, height: int
) -> list[str]:
    """把一页模板渲染成终端字符画。"""
    width = max(40, width)
    height = max(12, height)
    if not deck.slides:
        return ["（模板中没有页面）"]
    index = min(max(0, index), len(deck.slides) - 1)
    slide = deck.slides[index]
    canvas = _Canvas(width, height)
    canvas.rect(0, 0, width - 1, height - 1)
    title = slide.title or deck.title or "模板"
    canvas.put(2, 0, f" {_truncate(title, width - 6)} ")

    for block in slide.blocks:
        box = block.box
        if box is None:
            continue
        # 右上角页眉槽位在终端预览中不作为独立框显示。
        if box.row == 1 and box.col >= 9:
            line = _block_lines(block)[0] if _block_lines(block) else ""
            canvas.put(max(1, width - _display_width(line) - 2), 0, f" {line} ")
            continue
        x1, y1, x2, y2 = _map_box(box, width, height)
        if x2 - x1 < 2:
            continue
        if y2 - y1 < 3:
            # 终端行数不足时只画文字，避免 1–2 行高的块被直接跳过。
            compact_lines = _block_lines(block)
            compact_width = x2 - x1 + 1
            for offset, raw in enumerate(compact_lines[: max(1, y2 - y1 + 1)]):
                text = _truncate(raw, compact_width)
                text_width = _display_width(text)
                if box.align == "center":
                    horizontal = max(0, (compact_width - text_width) // 2)
                elif box.align == "right":
                    horizontal = max(0, compact_width - text_width)
                else:
                    horizontal = 0
                canvas.put(x1 + horizontal, y1 + offset, text, limit=x2 + 1)
            continue
        canvas.rect(x1, y1, x2, y2)
        inner_width = x2 - x1 - 1
        lines: list[str] = []
        for raw in _block_lines(block):
            lines.extend(_wrap(raw, inner_width))
        limit = max(0, y2 - y1 - 1)
        visible = lines[:limit]
        vertical = 0
        if box.align in ("center", "right") and len(visible) < limit:
            vertical = (limit - len(visible)) // 2
        for offset, line in enumerate(visible):
            text = _truncate(line, inner_width)
            text_width = _display_width(text)
            if box.align == "center":
                horizontal = max(0, (inner_width - text_width) // 2)
            elif box.align == "right":
                horizontal = max(0, inner_width - text_width)
            else:
                horizontal = 0
            canvas.put(
                x1 + 1 + horizontal,
                y1 + 1 + vertical + offset,
                text,
                limit=x2,
            )
        if len(lines) > limit and limit > 0:
            canvas.put(x2 - 1, y2 - 1, "…")
    return canvas.lines()


def _run_tui(screen, path: Path) -> None:
    import curses

    try:
        curses.curs_set(0)
    except curses.error:
        pass
    screen.nodelay(True)
    screen.keypad(True)
    index = 0
    deck: GridDeck | None = None
    error = ""
    try:
        mtime = path.stat().st_mtime_ns
    except OSError:
        mtime = 0
    dirty = True

    while True:
        key = screen.getch()
        if key in (ord("q"), 27):
            return
        if key in (curses.KEY_LEFT, ord("h")):
            index = max(0, index - 1)
        elif key in (curses.KEY_RIGHT, ord("l")):
            index += 1
        elif key == ord("r"):
            dirty = True

        try:
            current = path.stat().st_mtime_ns
        except OSError:
            current = mtime
        if current != mtime:
            mtime = current
            dirty = True

        if dirty:
            try:
                loaded = load_grid(path)
                deck = loaded
                error = ""
                if deck.slides:
                    index = min(index, len(deck.slides) - 1)
            except (GridSyntaxError, OSError) as exc:
                # 保存瞬间可能读到半写入内容；保留上一版预览，只提示错误。
                error = str(exc)
            dirty = False

        rows, cols = screen.getmaxyx()
        screen.erase()
        preview_w, preview_h, offset_x, offset_y = fit_preview_size(cols, rows)
        if deck is not None:
            lines = preview_lines(deck, index, preview_w, preview_h)
        else:
            lines = [f"模板解析失败：{error}"] if error else ["正在读取模板…"]
        hint = terminal_hint(cols, rows)
        for line_no, line in enumerate(lines[:preview_h]):
            try:
                screen.addstr(
                    offset_y + line_no,
                    offset_x,
                    _truncate(line, preview_w),
                )
            except curses.error:
                pass

        total = len(deck.slides) if deck is not None else 0
        status = (
            f"{path.name}  {min(index + 1, max(1, total))}/{max(1, total)}  "
            "保存后自动刷新  ←/→ 翻页  r 刷新  q 退出"
        )
        if error:
            status += f"  解析失败：{error}"
        if hint:
            status += f"  {hint}"
        try:
            screen.addstr(
                rows - 2,
                0,
                _truncate(status, max(0, cols - 1)),
                curses.A_REVERSE,
            )
        except curses.error:
            pass
        screen.refresh()
        time.sleep(0.15)


def run_template_adjust(template_dir: str | Path | None = None) -> int:
    """模板编写模式入口。"""
    paths = discover_templates(template_dir)
    if not paths:
        print(f"未找到模板：{Path(template_dir) if template_dir else DEFAULT_TEMPLATE_DIR}")
        return 2
    if not (sys.stdin.isatty() and sys.stdout.isatty()):
        print("模板编写模式需要交互终端。可用模板：")
        for path in paths:
            print(f"  - {path}")
        return 2
    try:
        selected = choose_template(paths)
    except (EOFError, KeyboardInterrupt):
        return 2
    cols, rows = shutil.get_terminal_size((80, 24))
    hint = terminal_hint(cols, rows)
    if hint:
        print(f"提示：{hint}")
    print(f"正在预览：{selected}（请用你的编辑器修改并保存）")
    try:
        import curses
    except ImportError:
        print("当前环境不支持 curses，无法提供终端预览。")
        return 2
    try:
        curses.wrapper(_run_tui, selected)
    except KeyboardInterrupt:
        pass
    except curses.error as exc:
        print(f"终端预览失败：{exc}")
        return 1
    return 0
