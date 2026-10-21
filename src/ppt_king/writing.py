"""逐页写作：检索资料证据，让模型输出带 16×9 坐标的 Markdown。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .cache import Cache
from .config import Settings
from .grid import (
    CODE,
    GRID_COLS,
    GRID_ROWS,
    IMAGE,
    TABLE,
    GridSyntaxError,
    box_anchor,
    parse_grid,
)
from .llm import LLMClient
from .outline import Outline, OutlinePage
from .sources import SourceBundle, SourceImage

PROMPT_VERSION = "write-v2"
MAX_EVIDENCE_CHARS = 7_000
MAX_EVIDENCE_PER_DOC = 5
MAX_IMAGES = 6
MAX_PARAGRAPH_CHARS = 1_200

_SPLIT_RE = re.compile(r"<!--\s*split\s*-->", re.IGNORECASE)
_PAGE_HEAD_RE = re.compile(r"^#{1,2}\s")
_STOP_WORDS = {
    "我们",
    "可以",
    "一个",
    "以及",
    "进行",
    "通过",
    "对于",
    "这个",
    "这些",
    "其中",
    "主要",
    "需要",
    "相关",
    "内容",
    "问题",
    "时候",
    "没有",
    "不是",
    "就是",
}


class WritingError(ValueError):
    """单页内容无法生成或不合规。"""


@dataclass
class PageResult:
    number: int
    title: str
    chunks: list[str] = field(default_factory=list)
    error: str = ""
    from_cache: bool = False

    @property
    def ok(self) -> bool:
        return not self.error and bool(self.chunks)


SYSTEM_PROMPT = """你是 PPT 内容撰写器。输入是一页幻灯片的大纲与资料证据，请输出带 16×9 网格坐标的 Markdown 块。

## 画布
16 列 × 9 行，坐标从 1 开始，`1,1` 是左上角。每格是正方形。
页标题由程序添加。默认正文从第 2 行开始；若载荷说明有 Logo，则正文从第 3 行开始。

## 只输出块
不要输出 `#` 或 `##` 页标题，不要重复页标题文字，不要代码围栏，不要解释。

## 语法
```
@ 左,上-右,下 [left|center|right] [字号pt]
块内容（Markdown 原文）
```
`@` 单独占一行，作用于紧随其后的块；左上角和右下角都包含在块内。
例如 `@ 1,6-16,8` 表示从第 1 列第 6 行到第 16 列第 8 行。
块可以是段落、`-` 列表、`1.` 列表、图片、表格、引用或公式。
公式块写成单独两行的 `$$ ... $$`，必须夹在两段文字之间，不要写进句子或列表。

## 行内格式
- 加粗：`**文字**`
- 红色强调：`<red>文字</red>`；每页至少标 1 处最关键结论或数字，最多 2 处
- 不要使用反引号包裹公式；公式直接写纯文本，例如 m·v·ω ≤ μmg

## 顶栏
左上角 8×1 是 Logo，不要放内容。右上角 8×1 是页眉槽位，可用一个
`@ 9,1-16,1 right 20` 的短文本块填写本页章节名或一句结论（不超过 20 字），
也可留空；不要在这里放正文或图片。

## 布局选择
1. 纯文字：`@ 1,2-16,8`
2. 文字 + 一张图：文字 `@ 1,2-8,8`，图片 `@ 10,3-16,7`
3. 文字 + 两张图：文字 `@ 1,2-7,8`，图片 `@ 9,2-16,5` 与 `@ 9,6-16,8`
4. 只有一张图：`@ 2,2-15,8`
5. 只有表格：`@ 1,2-16,8`
6. 文字 + 表格：文字 `@ 1,2-16,4`，表格 `@ 1,6-16,8`
7. 封面：主标题 `@ 1,4-16,4 center 34`，副标题/汇报人 `@ 1,6-16,6 center 18`；不要用列表，不要重复“封面”二字
8. 章节页：编号 `@ 1,3-16,4 center 60`，标题 `@ 1,5-16,5 center 28`

## 表格优先
当资料证据里有多组同构指标（如电压、电流、效率、纹波、功率）或参数对比时，
优先输出 GFM 表格，不要逐条写成列表：
```
@ 1,2-16,6
| 指标 | 数值 | 单位 |
| --- | --- | --- |
| 输入电压 | 24 | V |
| 输出电压 | 5.036 | V |
```
表格列数不超过 6 列；表格块高度至少 3 格；表格前后不要插空块。
同一张表只放在最相关的一页，不要在封面、总结页或其他页面重复。

## 硬约束
1. 只使用资料证据中的事实、数字、单位和结论，不编造。
2. 完整覆盖本页要点；每页 1-3 个块，内容少时不要硬凑。
3. 坐标必须在 1..16 列、1..9 行内，块之间不重叠，横竖至少留 1 格。
4. 文字块高度按内容给足：每 2 行文字约 1 格，宁可给大不要给小；若本页只有文字、没有图片或表格，文字块必须占满 16 列。
5. 图片块至少 5 格宽、4 格高，路径必须原样来自可用图片清单。
6. 表格高度 = `ceil((行数+1) / 2)` 格，至少 3 格。
7. 内容过多时不要缩小字号，用 `<!-- split -->` 单独一行分页。
8. 同类指标达到 3 项时，用表格呈现，而不是列表；但同一张表不要跨页重复。
9. 公式必须单独成块：`@ 1,5-16,5 center 24` 后接 `$$ ... $$`，前后各放一段文字。
10. 封面、章节页、总结页不要放表格，除非本页要点明确要求数据对比。
11. 正文、表格、图片的底边不要超过第 8 行底部，第 9 行留空作为安全边距。
12. 本页要点涉及架构、结构、流程或示意，且可用图片清单中有与主题直接相关的图片时，必须使用一张；图片题注必须与图片实际内容一致，不要为了配图而歪曲图片含义。
13. 只输出 Markdown 块与 `@` 行。"""

CORRECTION_TEMPLATE = """你上一次的输出无法使用：{error}

请重新输出这一页。要求：
- 每行 `@ 左,上-右,下` 之后紧跟一个块
- 坐标在 1..16 / 1..9 范围内，块之间不重叠
- 图片只能使用清单里的路径，不能写别的路径
- 不要输出页标题、代码围栏或解释"""


# --------------------------------------------------------------------------
# 证据检索
# --------------------------------------------------------------------------
def split_paragraphs(text: str) -> list[str]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks = [block.strip() for block in re.split(r"\n\s*\n+", normalized)]
    paragraphs: list[str] = []
    for block in blocks:
        if not block:
            continue
        if len(block) <= MAX_PARAGRAPH_CHARS:
            paragraphs.append(block)
            continue
        for start in range(0, len(block), MAX_PARAGRAPH_CHARS):
            chunk = block[start : start + MAX_PARAGRAPH_CHARS].strip()
            if chunk:
                paragraphs.append(chunk)
    return paragraphs


def page_keywords(page: OutlinePage) -> set[str]:
    raw = page.keywords_text
    keywords: set[str] = set()
    for token in re.findall(r"[A-Za-z][A-Za-z0-9_.%+-]+", raw):
        keywords.add(token.lower())
    for number in re.findall(r"\d+(?:\.\d+)?", raw):
        keywords.add(number)
    for chunk in re.findall(r"[\u4e00-\u9fff]{2,}", raw):
        for size in (4, 3, 2):
            for index in range(0, len(chunk) - size + 1):
                gram = chunk[index : index + size]
                if gram not in _STOP_WORDS:
                    keywords.add(gram)
    return keywords


def _score(text: str, keywords: set[str]) -> int:
    lowered = text.lower()
    return sum(len(keyword) for keyword in keywords if keyword in lowered)


def select_evidence(
    bundle: SourceBundle,
    page: OutlinePage,
    *,
    max_chars: int = MAX_EVIDENCE_CHARS,
) -> str:
    """按关键词从每份资料中挑选相关段落，控制总长度。"""
    keywords = page_keywords(page)
    picked: list[tuple[int, int, str, str]] = []

    for doc_index, doc in enumerate(bundle.docs):
        paragraphs = split_paragraphs(doc.text)
        if not paragraphs:
            continue
        scored = [(_score(paragraph, keywords), index) for index, paragraph in enumerate(paragraphs)]
        positive = [item for item in scored if item[0] > 0]
        chosen = sorted(positive, reverse=True)[:MAX_EVIDENCE_PER_DOC] if positive else scored[:3]
        chosen_indices = sorted(index for _, index in chosen)
        for index in chosen_indices:
            picked.append((doc_index, index, doc.rel, paragraphs[index]))

    picked.sort(key=lambda item: (item[0], item[1]))

    lines: list[str] = []
    used = 0
    current_doc = ""
    for _, _, rel, paragraph in picked:
        if rel != current_doc:
            header = f"【资料：{rel}】"
            if used + len(header) > max_chars:
                break
            lines.append(header)
            used += len(header)
            current_doc = rel
        remaining = max_chars - used
        if remaining <= 0:
            break
        if len(paragraph) > remaining:
            paragraph = paragraph[: max(0, remaining - 1)] + "…"
        lines.append(paragraph)
        used += len(paragraph)
    return "\n\n".join(lines)


def select_images(
    bundle: SourceBundle,
    page: OutlinePage,
    *,
    limit: int = MAX_IMAGES,
) -> list[SourceImage]:
    """只返回与本页主题相关的图片，避免为了配图而错用无关素材。"""
    if not bundle.images:
        return []
    keywords = page_keywords(page)
    scored = sorted(
        (
            (_score(f"{image.title} {image.caption}", keywords), image)
            for image in bundle.images
        ),
        key=lambda item: item[0],
        reverse=True,
    )
    return [image for score, image in scored if score > 0][:limit]


# --------------------------------------------------------------------------
# 输出清洗与校验
# --------------------------------------------------------------------------
def _strip_outer_fence(text: str) -> str:
    lines = text.splitlines()
    if not lines or not lines[0].strip().startswith("```"):
        return text
    language = lines[0].strip()[3:].strip().lower()
    if language not in ("", "markdown", "md"):
        return text
    lines = lines[1:]
    while lines and not lines[-1].strip():
        lines.pop()
    if lines and lines[-1].strip().startswith("```"):
        lines.pop()
    return "\n".join(lines)


def clean_output(text: str) -> str:
    cleaned = text.strip()
    cleaned = re.sub(r"\n\s*-{3,}\s*$", "", cleaned).strip()
    cleaned = _strip_outer_fence(cleaned)
    cleaned = re.sub(r"(?<!`)`([^`\n]+)`(?!`)", r"\1", cleaned)
    lines = cleaned.splitlines()
    while lines and _PAGE_HEAD_RE.match(lines[0].strip()):
        lines.pop(0)
        while lines and not lines[0].strip():
            lines.pop(0)
    cleaned = "\n".join(lines).strip()
    return re.sub(r"\n{3,}", "\n\n", cleaned)


def split_chunks(body: str) -> list[str]:
    return [part.strip() for part in _SPLIT_RE.split(body) if part.strip()]


def _normalize_image_path(value: str) -> str:
    path = value.replace("\\", "/").strip()
    while path.startswith("./"):
        path = path[2:]
    return path


def _box_rect(box) -> tuple[int, int, int, int]:
    width = box.colspan or (GRID_COLS - box.col + 1)
    height = box.rowspan or (GRID_ROWS - box.row + 1)
    return box.col, box.row, box.col + width, box.row + height


def validate_page(
    page_number: int,
    title: str,
    body: str,
    allowed_images: set[str],
    *,
    logo: bool = False,
) -> None:
    if not body.strip():
        raise WritingError(f"第 {page_number} 页输出为空")
    if re.search(r"^#{1,2}\s", body, flags=re.MULTILINE):
        raise WritingError(f"第 {page_number} 页包含页标题，应该只输出块")
    try:
        deck = parse_grid(f"## {title}\n\n{body}")
    except GridSyntaxError as exc:
        raise WritingError(f"第 {page_number} 页网格语法错误：{exc}") from exc

    blocks = deck.slides[0].blocks
    if not blocks:
        raise WritingError(f"第 {page_number} 页没有任何块")
    missing = [block.preview[:16] for block in blocks if not block.is_positioned]
    if missing:
        raise WritingError(
            f"第 {page_number} 页有 {len(missing)} 个块缺少坐标（如「{missing[0]}」）"
        )

    is_cover = page_number == 1 and title.strip().lower() in {"封面", "cover"}
    if logo and not is_cover:
        for block in blocks:
            box = block.box
            if box is None:
                continue
            header = box.row == 1 and box.col >= 9 and (box.colspan or 1) <= 8
            if box.row < 3 and not header:
                raise WritingError(
                    f"第 {page_number} 页正文与页眉/标题重叠："
                    f"`{box_anchor(box)}` 应移到第 3 行及以后"
                )

    for index, block in enumerate(blocks):
        if block.box is None:
            continue
        ax1, ay1, ax2, ay2 = _box_rect(block.box)
        for other in blocks[index + 1 :]:
            if other.box is None:
                continue
            bx1, by1, bx2, by2 = _box_rect(other.box)
            if min(ax2, bx2) > max(ax1, bx1) and min(ay2, by2) > max(ay1, by1):
                raise WritingError(
                    f"第 {page_number} 页块重叠："
                    f"`{box_anchor(block.box)}` 与 "
                    f"`{box_anchor(other.box)}`，请重新分配坐标"
                )

    for block in blocks:
        if block.kind != IMAGE:
            continue
        path = _normalize_image_path(block.path or "")
        if not path:
            raise WritingError(f"第 {page_number} 页有图片块缺少路径")
        if ".." in Path(path).parts:
            raise WritingError(f"第 {page_number} 页图片路径非法：{path}")
        if path not in allowed_images:
            raise WritingError(f"第 {page_number} 页引用了不在清单中的图片：{path}")


# --------------------------------------------------------------------------
# 提示词与生成
# --------------------------------------------------------------------------
def build_page_payload(
    page: OutlinePage,
    outline: Outline,
    total: int,
    evidence: str,
    images: list[SourceImage],
    *,
    language: str = "zh",
    logo: bool = False,
) -> str:
    layout_note = (
        "版式：Logo 占左上角 8×1 格，右上角 8×1 是页眉槽位（可用 "
        "`@ 9,1-16,1 right 20` 写短文本），页面标题在第 2 行，正文从第 3 行开始，"
        "内容底边不超过第 8 行。"
        if logo
        else "版式：页面标题在第 1 行，正文从第 2 行开始，内容底边不超过第 8 行。"
    )
    lines = [
        f"演示文稿：{outline.title}",
        f"页码：{page.number}/{total}",
        f"页标题：{page.title or '(无)'}",
        f"语言：{language}",
        layout_note,
        "",
        "本页要点：",
    ]
    lines.extend(f"- {point}" for point in page.points)
    if not page.points:
        lines.append("- 根据页标题撰写简洁内容")
    lines.extend(["", "资料证据：", evidence or "（没有匹配到资料段落，请只围绕页标题写最小内容）"])
    lines.append("")
    if images:
        lines.append("可用图片（只能原样使用这些路径）：")
        for image in images:
            caption = f"（{image.caption}）" if image.caption else ""
            lines.append(f"- {image.rel}{caption}")
    else:
        lines.append("可用图片：无，不要输出图片块")
    return "\n".join(lines)


def _dump(
    dump_dir: Path | None, number: int, payload: str, output: str
) -> None:
    if dump_dir is None:
        return
    dump_dir.mkdir(parents=True, exist_ok=True)
    (dump_dir / f"p{number:02d}_payload.txt").write_text(payload, encoding="utf-8")
    (dump_dir / f"p{number:02d}_output.md").write_text(output, encoding="utf-8")


def write_page(
    page: OutlinePage,
    outline: Outline,
    total: int,
    bundle: SourceBundle,
    *,
    settings: Settings,
    client: LLMClient,
    cache: Cache,
    language: str = "zh",
    logo: bool = False,
    dump_dir: Path | None = None,
    use_cache: bool = True,
) -> PageResult:
    """生成一页；失败时抛 :class:`WritingError`。"""
    evidence = select_evidence(bundle, page)
    images = select_images(bundle, page)
    allowed_images = {image.rel for image in images}
    payload = build_page_payload(
        page,
        outline,
        total,
        evidence,
        images,
        language=language,
        logo=logo,
    )

    key = Cache.key(
        "generate-page",
        PROMPT_VERSION,
        settings.model,
        f"{settings.temperature}",
        payload,
    )
    cached = cache.get(key) if use_cache else None
    if cached is not None:
        body = cached
        chunks = split_chunks(body)
        _dump(dump_dir, page.number, payload, body)
        return PageResult(
            number=page.number, title=page.title, chunks=chunks, from_cache=True
        )

    raw = client.complete(SYSTEM_PROMPT, payload)
    body = clean_output(raw)
    try:
        chunks = split_chunks(body)
        for chunk in chunks:
            validate_page(
                page.number, page.title, chunk, allowed_images, logo=logo
            )
        if not chunks:
            raise WritingError(f"第 {page.number} 页输出为空")
    except WritingError as exc:
        correction = CORRECTION_TEMPLATE.format(error=exc)
        raw = client.complete(SYSTEM_PROMPT, f"{payload}\n\n{correction}")
        body = clean_output(raw)
        chunks = split_chunks(body)
        try:
            for chunk in chunks:
                validate_page(
                    page.number, page.title, chunk, allowed_images, logo=logo
                )
            if not chunks:
                raise WritingError(f"第 {page.number} 页输出为空")
        except WritingError as retry_exc:
            raise WritingError(
                f"第 {page.number} 页「{page.title}」重试后仍不合规：{retry_exc}"
            ) from retry_exc

    _dump(dump_dir, page.number, payload, body)
    cache.put(key, body)
    return PageResult(number=page.number, title=page.title, chunks=chunks)


def placeholder_chunk(page: OutlinePage, *, logo: bool = False) -> str:
    if page.number == 1:
        subtitle = page.points[0] if page.points else "自动生成"
        return (
            f"@ 1,4-16,4 center 34\n{page.title}\n\n"
            f"@ 1,6-16,6 center 18\n{subtitle}"
        )
    points = page.points or ["内容待补充"]
    start_row = 3 if logo else 2
    height = 6 if logo else 7
    bottom = start_row + height - 1
    return f"@ 1,{start_row}-16,{bottom}\n" + "\n".join(
        f"- {point}" for point in points
    )


_ANCHOR_LINE_RE = re.compile(r"^@\s*\d")
_TABLE_SEP_RE = re.compile(r"^\|[\s:|-]+\|$")
_TABLE_PREFERENCE_RE = re.compile(r"(数据|指标|参数|对比|结果|测试|实测|性能)")
_IMAGE_BLOCK_RE = re.compile(
    r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)\s]+)(?:\s+\"[^\"]*\")?\)\s*$"
)


def _split_anchor_blocks(chunk: str) -> list[str]:
    blocks: list[str] = []
    current: list[str] = []
    for line in chunk.splitlines():
        if _ANCHOR_LINE_RE.match(line.strip()):
            if current:
                blocks.append("\n".join(current).strip())
            current = [line.rstrip()]
        elif current:
            current.append(line.rstrip())
        elif line.strip():
            current = [line.rstrip()]
    if current:
        blocks.append("\n".join(current).strip())
    return [block for block in blocks if block]


def _table_signature(block: str) -> tuple[tuple[str, ...], ...] | None:
    rows: list[tuple[str, ...]] = []
    for line in block.splitlines()[1:]:
        stripped = line.strip()
        if not stripped.startswith("|") or _TABLE_SEP_RE.match(stripped):
            continue
        cells = tuple(cell.strip() for cell in stripped.strip("|").split("|"))
        rows.append(cells)
    if len(rows) < 2:
        return None
    return tuple(rows)


def dedupe_duplicate_tables(
    outline: Outline, results: list[PageResult]
) -> list[PageResult]:
    """同一张表只保留在最相关的一页，避免模型跨页复读。"""
    blocks_per_result: list[list[list[str]]] = []
    occurrences: dict[tuple[tuple[str, ...], ...], list[tuple[int, int, int]]] = {}

    for result_index, result in enumerate(results):
        chunk_blocks: list[list[str]] = []
        for chunk_index, chunk in enumerate(result.chunks):
            blocks = _split_anchor_blocks(chunk)
            chunk_blocks.append(blocks)
            for block_index, block in enumerate(blocks):
                signature = _table_signature(block)
                if signature is not None:
                    occurrences.setdefault(signature, []).append(
                        (result_index, chunk_index, block_index)
                    )
        blocks_per_result.append(chunk_blocks)

    removals: set[tuple[int, int, int]] = set()
    for places in occurrences.values():
        if len(places) < 2:
            continue

        def keeper_key(place: tuple[int, int, int]) -> tuple[int, int]:
            result_index = place[0]
            title = results[result_index].title
            preferred = 1 if _TABLE_PREFERENCE_RE.search(title) else 0
            return (preferred, -results[result_index].number)

        keeper = max(places, key=keeper_key)
        removals.update(place for place in places if place != keeper)

    if not removals:
        return results

    deduped: list[PageResult] = []
    for result_index, result in enumerate(results):
        chunks: list[str] = []
        for chunk_index, blocks in enumerate(blocks_per_result[result_index]):
            kept = [
                block
                for block_index, block in enumerate(blocks)
                if (result_index, chunk_index, block_index) not in removals
            ]
            text = "\n\n".join(kept).strip()
            if text:
                chunks.append(text)
        deduped.append(
            PageResult(
                number=result.number,
                title=result.title,
                chunks=chunks,
                error=result.error,
                from_cache=result.from_cache,
            )
        )
    return deduped


def _image_signature(block: str) -> tuple[str, str] | None:
    lines = [line.strip() for line in block.splitlines()[1:] if line.strip()]
    if not lines:
        return None
    match = _IMAGE_BLOCK_RE.match(lines[0])
    if match is None:
        return None
    path = match.group("path").replace("\\", "/").lstrip("./")
    return path, match.group("alt").strip()


def _image_relevance(alt: str, path: str, page: OutlinePage) -> int:
    page_text = page.title + " " + " ".join(page.points)
    score = 0
    for chunk in re.findall(r"[\u4e00-\u9fff]{2,}", alt):
        for index in range(len(chunk) - 1):
            if chunk[index : index + 2] in page_text:
                score += 1
    for token in re.findall(r"[A-Za-z0-9]{3,}", f"{path} {alt}"):
        if token.lower() in page_text.lower():
            score += 1
    return score


def dedupe_duplicate_images(
    outline: Outline, results: list[PageResult]
) -> list[PageResult]:
    """同一张图只保留在与页面内容最相关的一页，避免跨页复读。"""
    blocks_per_result: list[list[list[str]]] = []
    occurrences: dict[str, list[tuple[int, int, int, str]]] = {}

    for result_index, result in enumerate(results):
        chunk_blocks: list[list[str]] = []
        for chunk_index, chunk in enumerate(result.chunks):
            blocks = _split_anchor_blocks(chunk)
            chunk_blocks.append(blocks)
            for block_index, block in enumerate(blocks):
                signature = _image_signature(block)
                if signature is not None:
                    path, alt = signature
                    occurrences.setdefault(path, []).append(
                        (result_index, chunk_index, block_index, alt)
                    )
        blocks_per_result.append(chunk_blocks)

    removals: set[tuple[int, int, int]] = set()
    for path, places in occurrences.items():
        if len(places) < 2:
            continue

        def keeper_key(
            place: tuple[int, int, int, str], image_path: str = path
        ) -> tuple[int, int]:
            result_index, _, _, alt = place
            page = outline.pages[result_index]
            relevance = _image_relevance(alt, image_path, page)
            return (relevance, -results[result_index].number)

        keeper = max(places, key=keeper_key)
        removals.update(
            place[:3] for place in places if place[:3] != keeper[:3]
        )

    if not removals:
        return results

    deduped: list[PageResult] = []
    for result_index, result in enumerate(results):
        chunks: list[str] = []
        for chunk_index, blocks in enumerate(blocks_per_result[result_index]):
            kept = [
                block
                for block_index, block in enumerate(blocks)
                if (result_index, chunk_index, block_index) not in removals
            ]
            text = "\n\n".join(kept).strip()
            if text:
                chunks.append(text)
        deduped.append(
            PageResult(
                number=result.number,
                title=result.title,
                chunks=chunks,
                error=result.error,
                from_cache=result.from_cache,
            )
        )
    return deduped


def assemble_deck(
    outline: Outline, results: list[PageResult], *, logo: bool = False
) -> str:
    results = dedupe_duplicate_tables(outline, results)
    results = dedupe_duplicate_images(outline, results)
    lines = [f"# {outline.title}", ""]
    for page, result in zip(outline.pages, results):
        chunks = result.chunks or [placeholder_chunk(page, logo=logo)]
        for index, chunk in enumerate(chunks):
            title = page.title or "未命名页"
            if index:
                title = f"{title}（续）"
            lines.extend([f"## {title}", "", chunk, ""])
    return "\n".join(lines).strip() + "\n"
