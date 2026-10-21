"""导出编排：PPTX -> IR -> LLM -> Markdown 文件。"""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field, replace
from pathlib import Path

from . import markdown as md
from .assets import AssetSink
from .cache import Cache
from .config import Settings, load_settings
from .ir import Deck, EmbedBlock, ImageBlock, Slide
from .llm import LLMClient, LLMError
from .prompt import SYSTEM_PROMPT, build_correction_message, build_user_message
from .reader import read_pptx


@dataclass
class ExportResult:
    source: Path
    output_dir: Path
    markdown_path: Path | None
    markdown: str
    deck: Deck
    pages: list[str] = field(default_factory=list)
    stats: str = ""


def _page_assets(slide: Slide) -> tuple[dict[str, ImageBlock], dict[str, EmbedBlock]]:
    images: dict[str, ImageBlock] = {}
    embeds: dict[str, EmbedBlock] = {}
    for block in slide.blocks:
        if isinstance(block, ImageBlock):
            images[block.asset_id] = block
        elif isinstance(block, EmbedBlock):
            embeds[block.asset_id] = block
    return images, embeds


def drop_repeated_images(deck: Deck, min_ratio: float = 0.6) -> int:
    """移除在多页重复出现的图片（通常是模板装饰/页眉logo），返回移除数量。"""
    counts: dict[str, int] = {}
    for image in deck.images:
        counts[image.rel_path] = counts.get(image.rel_path, 0) + 1
    threshold = max(2, int(len(deck.slides) * min_ratio))
    dropped = {path for path, count in counts.items() if count >= threshold}
    if not dropped:
        return 0
    removed = 0
    for slide in deck.slides:
        kept = [
            block
            for block in slide.blocks
            if not (isinstance(block, ImageBlock) and block.rel_path in dropped)
        ]
        removed += len(slide.blocks) - len(kept)
        slide.blocks = kept
    return removed


def _render_slide(
    client: LLMClient,
    cache: Cache,
    settings: Settings,
    slide: Slide,
    total: int,
    include_notes: bool,
    dump_dir: Path | None,
) -> str:
    payload = build_user_message(slide, total, include_notes=include_notes)
    images, embeds = _page_assets(slide)
    allowed = set(images) | set(embeds)

    key = Cache.key(
        settings.model,
        f"{settings.temperature}",
        f"{include_notes}",
        SYSTEM_PROMPT,
        payload,
    )

    cached = cache.get(key)
    if cached is not None:
        text = cached
        _dump(dump_dir, slide.number, payload, text, cached=True)
    else:
        text = md.normalize(client.complete(SYSTEM_PROMPT, payload))
        problems = md.validate_page(text, allowed)
        if problems:
            correction = build_correction_message(slide.number, problems)
            text = md.normalize(client.complete(SYSTEM_PROMPT, f"{payload}\n\n{correction}"))
            remaining = md.validate_page(text, allowed)
            if remaining:
                raise LLMError(
                    f"第 {slide.number} 页输出仍不合规：{'；'.join(remaining)}\n"
                    f"模型输出：\n{text}"
                )
        _dump(dump_dir, slide.number, payload, text, cached=False)
        cache.put(key, text)

    text = md.ensure_heading(text, slide.number, slide.title)
    text = md.append_missing(text, md.missing_asset_ids(text, slide))
    return md.resolve_placeholders(text, images, embeds)


def _prune_unused(output_dir: Path, document: str, sink: AssetSink) -> int:
    """删除最终 Markdown 未引用的资源（例如被过滤掉的重复图片）。"""
    removed = 0
    for path in sink.written:
        try:
            rel = path.relative_to(output_dir).as_posix()
        except ValueError:
            continue
        if rel in document:
            continue
        try:
            path.unlink()
            removed += 1
        except OSError:
            pass
    return removed


def _dump(dump_dir: Path | None, number: int, payload: str, output: str, cached: bool) -> None:
    if dump_dir is None:
        return
    dump_dir.mkdir(parents=True, exist_ok=True)
    (dump_dir / f"slide{number:03d}_payload.txt").write_text(payload, encoding="utf-8")
    marker = "（来自缓存）" if cached else ""
    (dump_dir / f"slide{number:03d}_output.md").write_text(
        f"<!-- {marker} -->\n{output}", encoding="utf-8"
    )


def export_pptx(
    path: str | Path,
    *,
    out_root: str | Path = "export",
    settings: Settings | None = None,
    include_notes: bool = True,
    use_cache: bool = True,
    dump_dir: str | Path | None = None,
    concurrency: int | None = None,
    write_file: bool = True,
    progress: bool = True,
    drop_repeated: bool = False,
) -> ExportResult:
    source = Path(path)
    settings = settings or load_settings()
    if concurrency:
        settings = replace(settings, concurrency=concurrency)

    output_dir = Path(out_root) / source.stem
    sink = AssetSink(root=output_dir)

    deck = read_pptx(source, sink)
    if drop_repeated:
        removed = drop_repeated_images(deck)
        if progress and removed:
            print(f"[ppt-king] 已忽略 {removed} 张多页重复图片", file=sys.stderr)
    total = len(deck.slides)
    if total == 0:
        raise LLMError(f"{source.name} 没有任何幻灯片")

    client = LLMClient(settings)
    cache = Cache.create(settings.cache_dir, enabled=use_cache)
    dump_path = Path(dump_dir) if dump_dir else None

    if progress:
        print(
            f"[ppt-king] {source.name}: {total} 页，{settings.model} @ {settings.base_url}",
            file=sys.stderr,
        )

    def run(slide: Slide) -> str:
        page = _render_slide(client, cache, settings, slide, total, include_notes, dump_path)
        if progress:
            print(f"[ppt-king] 第 {slide.number}/{total} 页完成", file=sys.stderr)
        return page

    workers = max(1, min(settings.concurrency, total))
    if workers == 1:
        pages = [run(slide) for slide in deck.slides]
    else:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            pages = list(pool.map(run, deck.slides))

    document = md.assemble(deck.title, pages)
    markdown_path: Path | None = None
    if write_file:
        output_dir.mkdir(parents=True, exist_ok=True)
        markdown_path = output_dir / f"{source.stem}.md"
        markdown_path.write_text(document, encoding="utf-8")
    _prune_unused(output_dir, document, sink)

    return ExportResult(
        source=source,
        output_dir=output_dir,
        markdown_path=markdown_path,
        markdown=document,
        deck=deck,
        pages=pages,
        stats=client.stats(),
    )
