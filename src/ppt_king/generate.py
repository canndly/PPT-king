"""生成编排：资料目录 -> 大纲 -> 逐页网格 Markdown -> PPTX。"""

from __future__ import annotations

import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field, replace
from pathlib import Path

from .cache import Cache
from .config import Settings, load_settings
from .grid import GridSyntaxError, load_grid
from .grid_render import make_theme, render_deck
from .llm import LLMClient, LLMError
from .outline import (
    MAX_PAGES,
    Outline,
    OutlineError,
    ensure_cover,
    generate_outline,
    outline_to_markdown,
    parse_outline,
    validate_outline,
)
from .sources import SourceBundle, describe_skipped, read_text_file, scan_sources
from .writing import (
    PageResult,
    WritingError,
    assemble_deck,
    dedupe_duplicate_images,
    dedupe_duplicate_tables,
    placeholder_chunk,
    split_chunks,
    write_page,
)


DEFAULT_OUT_ROOT = Path("out")


class GenerateError(RuntimeError):
    """生成流程失败。"""


@dataclass
class GenerateResult:
    output_dir: Path
    outline_path: Path
    grid_path: Path
    pptx_path: Path
    pages: int
    failures: list[str] = field(default_factory=list)
    cached_pages: int = 0
    stats: str = ""


def default_output_dir(source: str | Path) -> Path:
    """默认输出目录：项目当前目录下的 ``out/<资料目录名>``。"""
    return (DEFAULT_OUT_ROOT / Path(source).name).resolve()


def _copy_outline(source: Path, target: Path) -> str:
    text = read_text_file(source)
    if source.resolve() != target.resolve():
        target.write_text(text, encoding="utf-8")
    return text


def _prepare_outline(
    bundle: SourceBundle,
    out_dir: Path,
    *,
    settings: Settings,
    client: LLMClient,
    cache: Cache,
    min_pages: int,
    max_pages: int,
    language: str,
    resume: bool,
    force: bool,
    progress: bool,
) -> tuple[Outline, Path]:
    target = out_dir / "outline.md"
    target.parent.mkdir(parents=True, exist_ok=True)

    if bundle.outline:
        text = _copy_outline(bundle.outline, target)
        if progress:
            print(f"[ppt-king] 使用大纲：{bundle.outline}", file=sys.stderr)
    elif resume and target.is_file() and not force:
        text = target.read_text(encoding="utf-8-sig")
        if progress:
            print(f"[ppt-king] 复用大纲：{target}", file=sys.stderr)
    else:
        if progress:
            print("[ppt-king] 正在根据资料生成大纲…", file=sys.stderr)
        text = generate_outline(
            bundle,
            settings=settings,
            client=client,
            cache=cache,
            min_pages=min_pages,
            max_pages=max_pages,
            language=language,
            progress=progress,
        )

    outline = parse_outline(text, fallback_title=bundle.root.name)
    outline = ensure_cover(outline)
    problems = validate_outline(outline, min_pages=1, max_pages=MAX_PAGES)
    if problems:
        raise OutlineError("；".join(problems))
    target.write_text(outline_to_markdown(outline), encoding="utf-8")
    return outline, target


def _page_path(pages_dir: Path, number: int) -> Path:
    return pages_dir / f"p{number:02d}.md"


def _read_page(path: Path) -> list[str]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except OSError:
        return []
    return split_chunks(text)


def _write_page(path: Path, chunks: list[str]) -> None:
    body = "\n\n<!-- split -->\n\n".join(chunks)
    path.write_text(body.strip() + "\n", encoding="utf-8")


def _run_page(
    page,
    outline: Outline,
    bundle: SourceBundle,
    pages_dir: Path,
    *,
    settings: Settings,
    client: LLMClient,
    cache: Cache,
    language: str,
    logo: bool,
    dump_dir: Path | None,
    use_cache: bool,
    resume: bool,
    force: bool,
) -> PageResult:
    target = _page_path(pages_dir, page.number)
    if resume and not force and target.is_file():
        chunks = _read_page(target)
        if chunks:
            return PageResult(
                number=page.number,
                title=page.title,
                chunks=chunks,
                from_cache=True,
            )
    try:
        result = write_page(
            page,
            outline,
            len(outline.pages),
            bundle,
            settings=settings,
            client=client,
            cache=cache,
            language=language,
            logo=logo,
            dump_dir=dump_dir,
            use_cache=use_cache,
        )
    except (WritingError, LLMError) as exc:
        return PageResult(number=page.number, title=page.title, error=str(exc))

    _write_page(target, result.chunks)
    return result


def _write_manifest(
    out_dir: Path,
    *,
    source: Path,
    outline: Outline,
    settings: Settings,
    failures: list[str],
    bundle: SourceBundle,
    stats: str,
) -> None:
    payload = {
        "version": 1,
        "source": str(source),
        "outline": str(out_dir / "outline.md"),
        "grid": str(out_dir / "slides.grid.md"),
        "model": settings.model,
        "pages": len(outline.pages),
        "failures": failures,
        "skipped": bundle.skipped,
        "stats": stats,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    (out_dir / "generate.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def generate_deck(
    source: str | Path,
    *,
    out: str | Path | None = None,
    outline: str | Path | None = None,
    settings: Settings | None = None,
    resume: bool = False,
    force: bool = False,
    min_pages: int = 8,
    max_pages: int = 20,
    language: str = "zh",
    accent: str | None = None,
    logo: str | Path | bool | None = None,
    header: str | None = None,
    concurrency: int | None = None,
    use_cache: bool = True,
    dump_dir: str | Path | None = None,
    strict: bool = False,
    progress: bool = True,
) -> GenerateResult:
    """从资料目录生成 PPTX，中间产物全部落在输出目录。"""
    source_path = Path(source).resolve()
    if not source_path.is_dir():
        raise GenerateError(f"资料目录不存在：{source}")
    if min_pages < 1 or max_pages < min_pages:
        raise GenerateError("页数范围不合法")

    settings = settings or load_settings()
    if concurrency:
        settings = replace(settings, concurrency=concurrency)

    # 默认 Logo（None）同样要占用标题上方的页眉区，逐页写作必须从第 3 行开始。
    use_logo = logo is not False

    out_dir = Path(out).resolve() if out else default_output_dir(source_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    if progress:
        print(f"[ppt-king] 扫描资料：{source_path}", file=sys.stderr)
    bundle = scan_sources(source_path, out_dir=out_dir, outline=outline)
    skipped_note = describe_skipped(bundle)
    if progress and skipped_note:
        print(f"[ppt-king] {skipped_note}", file=sys.stderr)
    if progress:
        print(
            f"[ppt-king] 读取 {len(bundle.docs)} 份文本资料、{len(bundle.images)} 张图片",
            file=sys.stderr,
        )

    client = LLMClient(settings)
    cache = Cache.create(settings.cache_dir, enabled=use_cache)
    dump_path = Path(dump_dir) if dump_dir else None

    outline_obj, outline_path = _prepare_outline(
        bundle,
        out_dir,
        settings=settings,
        client=client,
        cache=cache,
        min_pages=min_pages,
        max_pages=max_pages,
        language=language,
        resume=resume,
        force=force,
        progress=progress,
    )

    grid_path = out_dir / "slides.grid.md"
    failures: list[str] = []
    cached_pages = 0

    if resume and grid_path.is_file() and not force:
        if progress:
            print(f"[ppt-king] 复用已有网格 Markdown：{grid_path}", file=sys.stderr)
        results: list[PageResult] = []
    else:
        pages_dir = out_dir / "pages"
        pages_dir.mkdir(parents=True, exist_ok=True)
        total = len(outline_obj.pages)

        def run(page) -> PageResult:
            result = _run_page(
                page,
                outline_obj,
                bundle,
                pages_dir,
                settings=settings,
                client=client,
                cache=cache,
                language=language,
                logo=use_logo,
                dump_dir=dump_path,
                use_cache=use_cache,
                resume=resume,
                force=force,
            )
            if progress:
                state = "缓存" if result.from_cache else ("失败" if result.error else "完成")
                print(f"[ppt-king] 第 {page.number}/{total} 页{state}", file=sys.stderr)
            return result

        workers = max(1, min(settings.concurrency, total))
        if workers == 1:
            results = [run(page) for page in outline_obj.pages]
        else:
            with ThreadPoolExecutor(max_workers=workers) as pool:
                results = list(pool.map(run, outline_obj.pages))

        results.sort(key=lambda item: item.number)
        results = dedupe_duplicate_tables(outline_obj, results)
        results = dedupe_duplicate_images(outline_obj, results)
        for result in results:
            if result.error:
                continue
            chunks = result.chunks or [
                placeholder_chunk(
                    outline_obj.pages[result.number - 1], logo=use_logo
                )
            ]
            _write_page(_page_path(pages_dir, result.number), chunks)

        failures = [
            f"P{result.number:02d}「{result.title}」：{result.error}"
            for result in results
            if result.error
        ]
        cached_pages = sum(1 for result in results if result.from_cache)

        grid_text = assemble_deck(outline_obj, results, logo=use_logo)
        grid_path.write_text(grid_text, encoding="utf-8")

    if failures and strict:
        _write_manifest(
            out_dir,
            source=source_path,
            outline=outline_obj,
            settings=settings,
            failures=failures,
            bundle=bundle,
            stats=client.stats(),
        )
        raise GenerateError("以下页面生成失败：\n" + "\n".join(failures))

    try:
        deck = load_grid(grid_path)
        pptx_path, report = render_deck(
            deck,
            out_dir / f"{source_path.name}.pptx",
            theme=make_theme(accent),
            base_dir=out_dir,
            logo=logo,
            header=header,
        )
    except (GridSyntaxError, ValueError) as exc:
        raise GenerateError(f"渲染失败：{exc}") from exc

    if progress:
        print(
            f"[ppt-king] {report.slides} 页 / {report.blocks} 块 -> {pptx_path}",
            file=sys.stderr,
        )
        for item in report.missing_images:
            print(
                f"[ppt-king] 图片不存在：{item}（已用占位文字代替）",
                file=sys.stderr,
            )
        for item in report.overflows[:10]:
            print(f"[ppt-king] 溢出警告：{item}", file=sys.stderr)

    _write_manifest(
        out_dir,
        source=source_path,
        outline=outline_obj,
        settings=settings,
        failures=failures,
        bundle=bundle,
        stats=client.stats(),
    )
    return GenerateResult(
        output_dir=out_dir,
        outline_path=outline_path,
        grid_path=grid_path,
        pptx_path=pptx_path,
        pages=len(outline_obj.pages),
        failures=failures,
        cached_pages=cached_pages,
        stats=client.stats(),
    )
