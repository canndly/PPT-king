"""大纲：解析人工大纲，或由模型从资料归纳。"""

from __future__ import annotations

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

from .cache import Cache
from .config import Settings
from .llm import LLMClient
from .sources import SourceBundle, SourceDoc

PROMPT_VERSION = "outline-v1"
MAX_POINTS_PER_PAGE = 8
MAX_POINT_CHARS = 160
MAX_DOC_CHARS = 8_000
MAX_SUMMARY_CHARS = 1_200
MAX_PAGES = 40
COVER_TITLES = ("封面", "封面页", "标题页", "cover", "title")

_PAGE_RE = re.compile(r"^##(?:\s+(.*))?$")
_DECK_RE = re.compile(r"^#(?:\s+(.*))?$")
_BULLET_RE = re.compile(r"^(?:[-*+•·]|\d+[.)、])\s+(.*)$")
_TITLE_INDEX_RE = re.compile(r"^\d{1,2}\s*[.、．)）]\s*")
_FENCE_RE = re.compile(r"^```[a-zA-Z]*\s*$")


class OutlineError(ValueError):
    """大纲无法解析或不符合约定。"""


@dataclass
class OutlinePage:
    number: int
    title: str
    points: list[str] = field(default_factory=list)

    @property
    def keywords_text(self) -> str:
        return " ".join([self.title, *self.points])


@dataclass
class Outline:
    title: str
    pages: list[OutlinePage] = field(default_factory=list)


SUMMARY_SYSTEM_PROMPT = """你是资料分析员。阅读用户提供的一份资料，输出供制作 PPT 使用的结构化摘要。

要求：
- 保留关键事实、数字、单位、结论、术语与因果关系。
- 用短句条目，最多 12 条；每条不超过 60 字。
- 不编造资料中没有的信息，不做空泛评价。
- 只输出摘要正文，不要标题、不要解释、不要代码围栏。"""

OUTLINE_SYSTEM_PROMPT = """你是 PPT 大纲策划。根据资料摘要，输出一份可直接用于逐页撰写幻灯片的大纲。

## 输出格式
# 演示文稿标题

## 页标题

- 要点一
- 要点二

## 下一页标题

- 要点一

## 硬约束
1. 严格按用户给定的页数范围安排页面。
2. 第一页是封面，最后一页是总结或致谢；中间页面按逻辑递进。
3. 页标题不超过 20 字；每页 2-5 条要点，每条不超过 60 字。
4. 只依据资料摘要，不编造事实、数字或结论。
5. 只输出大纲 Markdown，不要解释、不要代码围栏、不要使用 `###`。"""

OUTLINE_CORRECTION = """上一次输出的大纲不符合要求：{problems}

请重新输出，严格遵守：
- 第一行是 `# 演示文稿标题`
- 每页以 `## 页标题` 开头，下面跟 2-5 条 `- 要点`
- 页数必须在 {min_pages}..{max_pages} 之间
- 不要代码围栏，不要解释"""


# --------------------------------------------------------------------------
# 解析
# --------------------------------------------------------------------------
def _clean_text(value: str) -> str:
    text = value.strip()
    text = text.replace("**", "").replace("`", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _clean_title(value: str) -> str:
    text = _clean_text(value)
    cleaned = _TITLE_INDEX_RE.sub("", text, count=1).strip()
    return cleaned or text


def _collect_points(lines: list[str]) -> list[str]:
    points: list[str] = []
    for raw in lines:
        line = raw.strip()
        if not line or _FENCE_RE.match(line):
            continue
        bullet = _BULLET_RE.match(line)
        text = bullet.group(1) if bullet else line.lstrip("#").strip()
        text = _clean_text(text)
        if not text or text in points:
            continue
        points.append(text[:MAX_POINT_CHARS])
        if len(points) >= MAX_POINTS_PER_PAGE:
            break
    return points


def parse_outline(text: str, *, fallback_title: str = "演示文稿") -> Outline:
    """解析 ``# 标题`` + ``## 页标题`` + ``- 要点``；兼容纯列表大纲。"""
    cleaned = text.replace("\r\n", "\n").strip()
    lines = [line for line in cleaned.split("\n") if not _FENCE_RE.match(line.strip())]
    has_pages = any(_PAGE_RE.match(line) for line in lines)

    title = fallback_title
    pages: list[OutlinePage] = []

    if has_pages:
        current: OutlinePage | None = None
        buffer: list[str] = []

        def flush() -> None:
            nonlocal current, buffer
            if current is not None:
                current.points = _collect_points(buffer)
                pages.append(current)
            current = None
            buffer = []

        for line in lines:
            deck = _DECK_RE.match(line)
            if deck:
                candidate = _clean_text(deck.group(1) or "")
                if candidate and title == fallback_title:
                    title = candidate
                continue
            page = _PAGE_RE.match(line)
            if page:
                flush()
                current = OutlinePage(
                    number=len(pages) + 1,
                    title=_clean_title(page.group(1) or "") or f"第 {len(pages) + 1} 页",
                )
                continue
            if current is not None:
                buffer.append(line)
        flush()
    else:
        # 兼容「一页一个顶级列表项」的简化大纲
        current = None
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            bullet = _BULLET_RE.match(stripped)
            indent = len(line) - len(line.lstrip(" "))
            if bullet and indent == 0:
                if current is not None:
                    pages.append(current)
                current = OutlinePage(
                    number=len(pages) + 1,
                    title=_clean_title(bullet.group(1)) or f"第 {len(pages) + 1} 页",
                )
                continue
            if current is None:
                current = OutlinePage(
                    number=len(pages) + 1,
                    title=_clean_title(stripped) or fallback_title,
                )
                continue
            text_value = _clean_text(bullet.group(1) if bullet else stripped)
            if text_value and text_value not in current.points:
                current.points.append(text_value[:MAX_POINT_CHARS])
        if current is not None:
            pages.append(current)

    if not pages:
        raise OutlineError("大纲里没有任何页面，请使用 `## 页标题` 分页")
    return Outline(title=title, pages=pages)


def validate_outline(
    outline: Outline, *, min_pages: int = 1, max_pages: int = MAX_PAGES
) -> list[str]:
    problems: list[str] = []
    count = len(outline.pages)
    if count < min_pages:
        problems.append(f"页数 {count} 少于要求的 {min_pages}")
    if count > max_pages:
        problems.append(f"页数 {count} 多于要求的 {max_pages}")
    for page in outline.pages:
        if not page.title.strip():
            problems.append(f"第 {page.number} 页缺少标题")
        if len(page.points) > MAX_POINTS_PER_PAGE:
            page.points = page.points[:MAX_POINTS_PER_PAGE]
    return problems


def is_cover_title(title: str) -> bool:
    return title.strip().lower() in COVER_TITLES


def ensure_cover(outline: Outline, subtitle: str = "") -> Outline:
    """没有封面页时自动补一页，避免成品直接从正文开始。"""
    if outline.pages and is_cover_title(outline.pages[0].title):
        return outline
    points = [outline.title]
    if subtitle.strip():
        points.append(subtitle.strip())
    pages = [OutlinePage(number=1, title="封面", points=points)]
    for number, page in enumerate(outline.pages, 2):
        page.number = number
        pages.append(page)
    return Outline(title=outline.title, pages=pages)


def outline_to_markdown(outline: Outline) -> str:
    lines = [f"# {outline.title}", ""]
    for page in outline.pages:
        lines.append(f"## {page.title}")
        lines.append("")
        for point in page.points:
            lines.append(f"- {point}")
        if page.points:
            lines.append("")
    return "\n".join(lines).strip() + "\n"


def load_outline(path: str | Path, *, fallback_title: str = "演示文稿") -> Outline:
    target = Path(path)
    if not target.is_file():
        raise OutlineError(f"找不到大纲文件：{target}")
    return parse_outline(
        target.read_text(encoding="utf-8-sig"), fallback_title=fallback_title
    )


# --------------------------------------------------------------------------
# AI 归纳
# --------------------------------------------------------------------------
def _summarize(
    doc: SourceDoc,
    client: LLMClient,
    cache: Cache,
    settings: Settings,
) -> str:
    excerpt = doc.text[:MAX_DOC_CHARS]
    key = Cache.key(
        "generate-summary",
        PROMPT_VERSION,
        settings.model,
        f"{settings.temperature}",
        doc.rel,
        excerpt,
    )
    cached = cache.get(key)
    if cached is not None:
        return cached

    payload = f"资料：{doc.rel}\n标题：{doc.title}\n\n{excerpt}"
    summary = client.complete(SUMMARY_SYSTEM_PROMPT, payload).strip()
    summary = summary[:MAX_SUMMARY_CHARS]
    cache.put(key, summary)
    return summary


def _clean_outline_output(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        lines = cleaned.splitlines()[1:]
        while lines and not lines[-1].strip():
            lines.pop()
        if lines and lines[-1].strip().startswith("```"):
            lines.pop()
        cleaned = "\n".join(lines).strip()
    start = cleaned.find("# ")
    if start > 0:
        cleaned = cleaned[start:]
    return cleaned.strip()


def generate_outline(
    bundle: SourceBundle,
    *,
    settings: Settings,
    client: LLMClient | None = None,
    cache: Cache | None = None,
    min_pages: int = 8,
    max_pages: int = 20,
    language: str = "zh",
    progress: bool = True,
) -> str:
    """根据资料摘要生成大纲 Markdown；失败时抛 :class:`OutlineError`。"""
    if min_pages < 1:
        raise OutlineError("页数下限必须 >= 1")
    if max_pages < min_pages:
        raise OutlineError("页数上限不能小于下限")

    llm = client or LLMClient(settings)
    cache_store = cache or Cache.create(settings.cache_dir, enabled=True)

    workers = max(1, min(settings.concurrency, len(bundle.docs)))
    if workers == 1:
        summaries = [_summarize(doc, llm, cache_store, settings) for doc in bundle.docs]
    else:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            summaries = list(
                pool.map(lambda doc: _summarize(doc, llm, cache_store, settings), bundle.docs)
            )

    if progress:
        print(f"[ppt-king] 已摘要 {len(bundle.docs)} 份资料", file=sys.stderr)

    blocks = []
    for doc, summary in zip(bundle.docs, summaries):
        blocks.append(f"### 资料：{doc.rel}（{doc.title}）\n{summary}")

    payload = (
        f"资料目录：{bundle.root.name}\n"
        f"页数：{min_pages}..{max_pages}\n"
        f"语言：{language}\n\n"
        "资料摘要：\n" + "\n\n".join(blocks)
    )

    key = Cache.key(
        "generate-outline",
        PROMPT_VERSION,
        settings.model,
        f"{settings.temperature}",
        f"{min_pages}-{max_pages}",
        language,
        "\n\n".join(summaries),
    )
    cached = cache_store.get(key)
    if cached is not None:
        return cached

    raw = _clean_outline_output(llm.complete(OUTLINE_SYSTEM_PROMPT, payload))
    try:
        outline = parse_outline(raw, fallback_title=bundle.root.name)
        problems = validate_outline(outline, min_pages=min_pages, max_pages=max_pages)
        if problems:
            raise OutlineError("；".join(problems))
    except OutlineError as exc:
        correction = OUTLINE_CORRECTION.format(
            problems=str(exc), min_pages=min_pages, max_pages=max_pages
        )
        raw = _clean_outline_output(
            llm.complete(OUTLINE_SYSTEM_PROMPT, f"{payload}\n\n{correction}")
        )
        outline = parse_outline(raw, fallback_title=bundle.root.name)
        problems = validate_outline(outline, min_pages=min_pages, max_pages=max_pages)
        if problems:
            raise OutlineError("模型输出仍不合规：" + "；".join(problems)) from exc

    text = outline_to_markdown(outline)
    cache_store.put(key, text)
    return text
