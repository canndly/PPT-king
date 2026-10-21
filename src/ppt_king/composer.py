"""Markdown + 图片 -> 版式规格（DeckSpec），由大模型决定结构。"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image

from .cache import Cache
from .config import Settings
from .llm import LLMClient, LLMError
from .spec import DeckSpec, SlideSpec, SpecError, parse_slide

IMAGE_RE = re.compile(r"!\[(?P<alt>[^\]]*)\]\((?P<path>[^)\s]+)(?:\s+\"[^\"]*\")?\)")
HEADING_RE = re.compile(r"^##\s+(?P<rest>.*)$")
PAGE_NO_RE = re.compile(r"^(?P<num>\d+)\s*[.、．]\s*(?P<title>.*)$")
NOTE_RE = re.compile(r"^>\s*\*\*备注\*\*\s*[：:]\s*(?P<first>.*)$")

BUILD_PROMPT_VERSION = "v1"

SYSTEM_PROMPT = """你是 PPT 版式设计器。输入是一页幻灯片导出的 Markdown，请把它还原成 PPT 版式 JSON。

## 输出
只输出一个 JSON 对象，不要代码围栏，不要解释，不要前后缀。

{
  "layout": "title",
  "subtitle": "",
  "blocks": [],
  "notes": ""
}

不要输出 title 字段（页标题由程序负责）。

## layout 取值
- title：封面（大标题 + 副标题/姓名学号等短信息）
- section：章节过渡页（只有标题，没有正文内容）
- two_column：左侧文字、右侧图片
- image_left：左侧图片、右侧文字
- gallery：以图片为主（≥3 张独立图片）
- table：以表格为主
- title_content：以上都不适用时（单列纵向排列）

## block 类型
- {"type": "heading", "text": "小节标题"}
- {"type": "paragraph", "text": "正文段落"}
- {"type": "bullets", "items": [{"text": "条目", "level": 0}]}
- {"type": "image", "ref": "A1", "caption": "图注"}
- {"type": "table", "header": ["列1", "列2"], "rows": [["a", "b"]]}
- {"type": "quote", "text": "需要强调的说明或引用"}

## 规则
1. 图片只能写成 {"type": "image", "ref": "A1"}，ref 必须来自“本页可用图片”列表；
   禁止写路径、禁止编造 ref。原 Markdown 中的 {{A1}} 占位符必须各自变成一个 image 块，且不得遗漏。
2. `###` 级标题转为 heading 块；`####` 及更深合并进 heading 或 paragraph。
3. 图表小节（`### 图表：…` 后跟数据表）输出为 table 块，表头第一列为“分类”或原表头。
4. 去掉所有 Markdown 标记（`**`、`*`、`` ` ``、`#`、`|`、`---`），只保留文字内容。
5. 原文的 `<br>` 表示换行：同一语义组拆成相邻的 paragraph 或 bullets 条目。
6. 内容较多（≥3 个并列条目）时用 bullets；连续叙述用 paragraph。不改变原文措辞，不概括。
7. 页面里出现“6. 波形展示”这类用序号手写的标题，可保留为 paragraph 或 heading。
8. 单页元素很少（只有标题 + 一句副标题）时用 title 或 section。
9. notes 字段放演讲者备注；输入中的 `> **备注**：…` 已由程序剥离，通常留空字符串。
10. 不要输出任何 Markdown 或 HTML 标记。"""


@dataclass
class PageSource:
    number: int
    title: str
    body: str = ""
    notes: str = ""
    assets: dict[str, str] = field(default_factory=dict)  # ref -> 相对路径
    captions: dict[str, str] = field(default_factory=dict)  # ref -> alt

    @property
    def allowed_refs(self) -> set[str]:
        return set(self.assets)


@dataclass
class SplitDeck:
    title: str
    pages: list[PageSource]


def _split_heading(raw: str, fallback_number: int) -> tuple[int, str]:
    text = raw.strip()
    match = PAGE_NO_RE.match(text)
    if match:
        try:
            return int(match.group("num")), match.group("title").strip()
        except ValueError:
            pass
    match = re.match(r"^幻灯片\s*(\d+)\s*(.*)$", text)
    if match:
        return int(match.group(1)), match.group(2).strip()
    return fallback_number, text


def _extract_notes(lines: list[str]) -> tuple[list[str], str]:
    kept: list[str] = []
    notes: list[str] = []
    collecting = False
    for line in lines:
        match = NOTE_RE.match(line.strip())
        if match:
            collecting = True
            if match.group("first").strip():
                notes.append(match.group("first").strip())
            continue
        if collecting:
            stripped = line.strip()
            if stripped.startswith(">"):
                notes.append(stripped.lstrip(">").strip())
                continue
            if not stripped:
                collecting = False
                kept.append(line)
                continue
            collecting = False
        kept.append(line)
    return kept, "\n".join(note for note in notes if note).strip()


def _index_images(
    body: str, base_dir: Path
) -> tuple[str, dict[str, str], dict[str, str]]:
    """把 `![](assets/x.png)` 替换成 `{{A1}}`，返回 (新正文, ref->相对路径, ref->alt)。"""
    assets: dict[str, str] = {}
    captions: dict[str, str] = {}

    def _replace(match: re.Match[str]) -> str:
        raw_path = match.group("path").strip()
        alt = match.group("alt").strip()
        rel = raw_path.replace("\\", "/")
        if rel.startswith("./"):
            rel = rel[2:]
        if not (base_dir / rel).is_file():
            return match.group(0)  # 文件不存在，保留原文交给模型当普通文本
        ref = f"A{len(assets) + 1}"
        assets[ref] = rel
        captions[ref] = alt
        return "{{" + ref + "}}"

    return IMAGE_RE.sub(_replace, body), assets, captions


def split_markdown(text: str, base_dir: Path) -> SplitDeck:
    """按 `## ` 切页；`# ` 作为演示文稿标题。"""
    deck_title = base_dir.name
    lines = text.replace("\r\n", "\n").split("\n")

    pages: list[PageSource] = []
    current: PageSource | None = None
    buffer: list[str] = []

    def flush() -> None:
        nonlocal current, buffer
        if current is None:
            return
        raw = "\n".join(buffer)
        stripped, notes = _extract_notes(raw.split("\n"))
        body, assets, captions = _index_images("\n".join(stripped), base_dir)
        current.body = body.strip()
        current.notes = notes
        current.assets = assets
        current.captions = captions
        pages.append(current)
        current = None
        buffer = []

    for line in lines:
        if deck_title == base_dir.name and line.startswith("# "):
            candidate = line[2:].strip()
            if candidate:
                deck_title = candidate
            continue
        heading = HEADING_RE.match(line)
        if heading:
            flush()
            number, title = _split_heading(heading.group("rest"), len(pages) + 1)
            current = PageSource(number=number, title=title)
            continue
        if current is not None:
            buffer.append(line)

    flush()
    if not pages:
        raise SpecError(f"{base_dir} 中的 Markdown 没有找到 `## ` 页标题")
    return SplitDeck(title=deck_title, pages=pages)


def _image_note(ref: str, rel: str, captions: dict[str, str], base_dir: Path) -> str:
    caption = captions.get(ref, "")
    size = ""
    target = base_dir / rel
    try:
        with Image.open(target) as image:
            size = f"{image.width}x{image.height}"
    except Exception:  # pragma: no cover - 图片损坏时忽略尺寸
        size = ""
    parts = [f"{ref} = {rel}"]
    if caption:
        parts.append(f"图注={caption}")
    if size:
        parts.append(f"像素={size}")
    return "  ".join(parts)


def build_user_message(page: PageSource, total: int, base_dir: Path) -> str:
    lines = [f"幻灯片 {page.number}/{total}", f"页标题: {page.title or '(无)'}", "", page.body]
    if page.assets:
        lines.append("")
        lines.append("本页可用图片（只能引用这些 ref）：")
        for ref, rel in page.assets.items():
            lines.append("  " + _image_note(ref, rel, page.captions, base_dir))
    else:
        lines.append("")
        lines.append("本页可用图片：无")
    return "\n".join(lines)


def build_correction_message(number: int, problems: list[str]) -> str:
    bullets = "\n".join(f"- {problem}" for problem in problems)
    return (
        f"你上一次的输出不可用，请重新输出第 {number} 页的 JSON。\n"
        f"问题：\n{bullets}\n"
        "提醒：只输出一个 JSON 对象；layout 必须是给定取值之一；"
        "image 块的 ref 只能来自本页可用图片；table 必须给出 header。"
    )


def compose_slide(
    client: LLMClient,
    cache: Cache,
    settings: Settings,
    page: PageSource,
    total: int,
    base_dir: Path,
) -> SlideSpec:
    payload = build_user_message(page, total, base_dir)
    key = Cache.key(
        "build",
        BUILD_PROMPT_VERSION,
        settings.model,
        f"{settings.temperature}",
        SYSTEM_PROMPT,
        payload,
    )

    def parse(raw: str) -> SlideSpec:
        return parse_slide(
            raw,
            number=page.number,
            allowed_refs=page.allowed_refs,
            assets=page.assets,
            fallback_title=page.title,
        )

    cached = cache.get(key)
    if cached is not None:
        return parse(cached)

    raw = client.complete(SYSTEM_PROMPT, payload)
    try:
        slide = parse(raw)
    except SpecError as exc:
        correction = build_correction_message(page.number, [str(exc)])
        raw = client.complete(SYSTEM_PROMPT, f"{payload}\n\n{correction}")
        try:
            slide = parse(raw)
        except SpecError as retry_exc:
            raise LLMError(
                f"第 {page.number} 页版式 JSON 仍不合规：{retry_exc}\n模型输出：\n{raw}"
            ) from retry_exc

    cache.put(key, raw)
    return slide


def compose_deck(
    markdown_path: str | Path,
    *,
    settings: Settings,
    client: LLMClient,
    cache: Cache,
    concurrency: int = 4,
    progress: bool = True,
    dump_dir: Path | None = None,
) -> DeckSpec:
    md_path = Path(markdown_path)
    base_dir = md_path.parent
    split = split_markdown(md_path.read_text(encoding="utf-8"), base_dir)
    total = len(split.pages)

    def run(page: PageSource) -> SlideSpec:
        slide = compose_slide(client, cache, settings, page, total, base_dir)
        if dump_dir is not None:
            dump_dir.mkdir(parents=True, exist_ok=True)
            (dump_dir / f"slide{page.number:03d}.json").write_text(
                json.dumps(slide.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
            )
        if progress:
            print(f"[ppt-king] 版式 {page.number}/{total} 完成", file=sys.stderr)
        return slide

    workers = max(1, min(concurrency, total))
    if workers == 1:
        slides = [run(page) for page in split.pages]
    else:
        from concurrent.futures import ThreadPoolExecutor

        with ThreadPoolExecutor(max_workers=workers) as pool:
            slides = list(pool.map(run, split.pages))

    return DeckSpec(title=split.title, slides=slides, source=str(md_path))
