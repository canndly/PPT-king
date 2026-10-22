"""让模型给普通 Markdown 加上 16×9 网格坐标（逐页调用）。

分页策略与导出方向一致：单页载荷小、失败可重试、可并发。
模型只输出**该页的块与坐标**，页标题行由程序统一拼接，避免重复与格式漂移。
"""

from __future__ import annotations

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

from .cache import Cache
from .config import Settings
from .grid import GridSyntaxError, parse_grid
from .llm import LLMClient, LLMError

PROMPT_VERSION = "grid-v11"

DECK_TITLE_RE = re.compile(r"^#\s+(.*)$")
PAGE_HEAD_RE = re.compile(r"^##(?:\s+(.*))?$")

SYSTEM_PROMPT = """你是 PPT 网格排版器。输入是一页幻灯片的 Markdown 内容，请输出带 16×9 网格坐标的 Markdown 块。

## 画布
16 列 × 9 行，坐标从 1 开始，`1,1` 是左上角。每格是正方形。

## 你只输出块，不要输出页标题
页标题由程序添加在页面顶部（占第 1 行），默认左上角有 Logo。所以**正文从第 3 行开始**，
内容块左右各留 1 列（只放在第 2..15 列），并且**不要重复输出页标题文字**。

## 语法
```
@ 左,上-右,下 [left|center|right] [字号pt]
块内容（Markdown 原文）
```
`@` 单独占一行，作用于紧随其后的那一个块；左上角和右下角都包含在块内。
例如 `@ 1,6-16,8` 表示从第 1 列第 6 行到第 16 列第 8 行。对齐和字号可选。

## 块类型（由 Markdown 形态决定，不要改写内容）
- 段落：普通文字
- 列表：`-` 或 `1.`（引导句可以和列表写在同一个块里）
- 图片：`![说明](路径)`
- 表格：`| ... |`
- 引用：`> ...`

## 布局选择（只有这几种，不要发明别的）
1. 纯文字：一个块 `@ 2,3-15,8`
2. 文字 + 一张图：文字 `@ 2,3-8,8`，图片 `@ 10,3-15,7`
3. 文字 + 两张图：文字 `@ 2,3-7,8`，图片 `@ 9,3-15,4` 与 `@ 9,6-15,8`
4. 只有一张图：`@ 2,3-15,8`
5. 只有表格：`@ 2,3-15,8`
6. 文字 + 表格：文字 `@ 2,3-15,4`，表格 `@ 2,6-15,8`
7. 封面（只有短标题/副标题/日期）：每行一个块，居中
   `@ 2,4-15,4 center 34` / `@ 2,6-15,6 center 18` / `@ 2,8-15,8 center 12`
8. 章节页（只有一个大编号 + 一句标题）：编号 `@ 2,3-15,4 center 60`，标题 `@ 2,6-15,6 center 28`
9. 只有恰好 4 个同权重并列要点、且本页没有公式/图片/表格时，才每条单独成块并 2×2 定位 `@ 2,3-7,4`、`@ 10,3-15,4`、`@ 2,6-7,7`、`@ 10,6-15,7`；其他情况不要用 2×2
10. 结论/局限：左右两栏纯文字，`@ 2,3-7,8` 与 `@ 10,3-15,8`，每栏 2 条；不要加背景卡片或装饰
11. 含公式页面：说明文字（可包含结论）`@ 2,3-15,5`，公式 `@ 2,7-15,7 center 24`；固定 2 块纵向排列，不要单独放 1 行结论，不要 2×2
12. 总结页：只用 1 个整宽文字块 `@ 2,3-15,8`，不要拆成多块，不要 2×2

## 硬约束
1. 坐标必须落在 1..16 列、1..9 行内；宽×高不得越界。
2. 内容块左右各留 1 列（第 2..15 列），块之间横竖至少留 1 格间距，不得重叠。
3. 文本块高度按内容给足：每 2 行文字约 1 格；宁可给大不要给小。只有文字、没有图片或表格时，文本块使用 `@ 2,3-15,8`。
4. 图片块至少 4 格高、5 格宽。
5. 表格高度 = `ceil((行数+1) / 2)` 格，至少 3 格。
6. 一张图时就别硬塞成两栏（会显得空）。
7. 内容很多时**不要缩小字号**，而是把一页拆成两页：输出 `<!-- split -->` 单独一行，
   程序会在该处切分，后半部分成为新的一页（页标题会自动加"（续）"）。
8. 区块内容严格保真：不改写、不概括、不翻译、不合并句子。
9. 图片题注只能写在图片块内；不要单独输出没有对应图片的题注或说明文字。
10. 只输出 Markdown 块与 `@` 行，不要代码围栏，不要任何解释。"""

CORRECTION_TEMPLATE = """你上一次的输出无法解析：{error}

请重新输出。要求：
- 每行 `@ 左,上-右,下` 之后紧跟一个块，块内不要出现 `@` 行
- 内容块左右各留 1 列（第 2..15 列），块之间横竖至少留 1 格
- 含公式页面固定 2 块：说明（含结论）在上、公式在下，禁止 2×2
- 总结页只用一个整宽文字块，不要拆成多块
- 只有恰好 4 个同权重要点且没有公式/图片/表格时，才用 2×2
- 不要输出页标题，不要代码围栏，不要解释"""


@dataclass
class SourcePage:
    number: int
    title: str
    body: str = ""


@dataclass
class DeckPlan:
    title: str
    header: list[str] = field(default_factory=list)  # `# 标题` 之前的行
    pages: list[SourcePage] = field(default_factory=list)


def split_pages(text: str, *, fallback_title: str = "演示文稿") -> DeckPlan:
    """按 `## ` 切页，`# ` 作为演示文稿标题。"""
    plan = DeckPlan(title=fallback_title)
    current: SourcePage | None = None
    buffer: list[str] = []

    def flush() -> None:
        nonlocal current, buffer
        if current is not None:
            current.body = "\n".join(buffer).strip()
            plan.pages.append(current)
        current = None
        buffer = []

    for raw in text.replace("\r\n", "\n").split("\n"):
        page = PAGE_HEAD_RE.match(raw)
        if page:
            flush()
            current = SourcePage(number=len(plan.pages) + 1, title=(page.group(1) or "").strip())
            continue
        deck = DECK_TITLE_RE.match(raw)
        if deck:
            flush()
            title = deck.group(1).strip()
            if title:
                plan.title = title
            continue
        if current is None:
            if raw.strip():
                plan.header.append(raw)
            continue
        buffer.append(raw)

    flush()
    return plan


def _clean(text: str) -> str:
    cleaned = text.strip()
    cleaned = re.sub(r"^```[a-zA-Z]*\s*\n", "", cleaned)
    cleaned = re.sub(r"\n```\s*$", "", cleaned)
    # 模型偶尔会自己补页标题，去掉
    cleaned = re.sub(r"^#{1,2}\s+.*\n", "", cleaned, count=1) if re.match(
        r"^#{1,2}\s", cleaned
    ) else cleaned
    return cleaned.strip()


def _split_output(body: str) -> list[str]:
    parts = [part.strip() for part in body.split("<!-- split -->")]
    return [part for part in parts if part]


def _validate(page_no: int, body: str, title: str) -> None:
    """用完整语法校验，并强制每个块都带 `@` 坐标。"""
    stub = f"## {title}\n\n{body}"
    try:
        deck = parse_grid(stub)
    except GridSyntaxError as exc:
        raise GridSyntaxError(f"第 {page_no} 页：{exc}") from exc

    blocks = deck.slides[0].blocks
    if not blocks:
        raise GridSyntaxError(f"第 {page_no} 页：没有任何块")
    missing = [block.preview[:16] for block in blocks if not block.is_positioned]
    if missing:
        raise GridSyntaxError(
            f"第 {page_no} 页：有 {len(missing)} 个块缺少 `@ 左,上-右,下` 坐标"
            f"（如「{missing[0]}」）"
        )


def plan_page(
    client: LLMClient,
    cache: Cache,
    settings: Settings,
    page: SourcePage,
    total: int,
) -> list[str]:
    """返回该页的网格块文本（可能因 `<!-- split -->` 变成多段）。"""
    payload = f"页码 {page.number}/{total}\n页标题: {page.title or '(无)'}\n\n{page.body}"
    key = Cache.key(
        "grid-layout",
        PROMPT_VERSION,
        settings.model,
        f"{settings.temperature}",
        SYSTEM_PROMPT,
        payload,
    )

    cached = cache.get(key)
    if cached is not None:
        return _split_output(cached)

    raw = _clean(client.complete(SYSTEM_PROMPT, payload))
    try:
        _validate(page.number, raw, page.title)
    except GridSyntaxError as exc:
        raw = _clean(
            client.complete(SYSTEM_PROMPT, f"{payload}\n\n{CORRECTION_TEMPLATE.format(error=exc)}")
        )
        _validate(page.number, raw, page.title)

    cache.put(key, raw)
    return _split_output(raw)


def layout_markdown(
    path: str | Path,
    *,
    settings: Settings,
    use_cache: bool = True,
    client: LLMClient | None = None,
    concurrency: int | None = None,
    progress: bool = True,
) -> str:
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"找不到文件：{source}")

    plan = split_pages(source.read_text(encoding="utf-8-sig"), fallback_title=source.stem)
    if not plan.pages:
        raise GridSyntaxError("没有找到任何 `## 页标题`，无法确定页面")

    cache = Cache.create(settings.cache_dir, enabled=use_cache)
    llm = client or LLMClient(settings)
    total = len(plan.pages)

    def run(page: SourcePage) -> list[str]:
        chunks = plan_page(llm, cache, settings, page, total)
        if progress:
            print(f"[ppt-king] 排版 {page.number}/{total} 完成", file=sys.stderr)
        return chunks

    workers = max(1, min(concurrency or settings.concurrency, total))
    if workers == 1:
        results = [run(page) for page in plan.pages]
    else:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            results = list(pool.map(run, plan.pages))

    lines: list[str] = [f"# {plan.title}", ""]
    for page, chunks in zip(plan.pages, results):
        chunks = chunks or [""]
        for index, chunk in enumerate(chunks):
            title = page.title or ""
            if index:
                title = f"{title}（续）" if title else "（续）"
            lines.append(f"## {title}".rstrip())
            lines.append("")
            if chunk:
                lines.append(chunk)
                lines.append("")
    return "\n".join(lines).strip() + "\n"
