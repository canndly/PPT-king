"""还原编排：导出的 Markdown + 图片 -> DeckSpec -> PPTX。"""

from __future__ import annotations

import sys
from dataclasses import dataclass, replace
from pathlib import Path

from pptx.dml.color import RGBColor

from . import builder
from .builder import DEFAULT_THEME, Theme
from .cache import Cache
from .composer import compose_deck
from .config import Settings, load_settings
from .llm import LLMClient
from .spec import DeckSpec, SpecError


@dataclass
class RebuildResult:
    markdown: Path | None
    output: Path
    spec: DeckSpec
    spec_path: Path | None = None
    stats: str = ""


def locate_markdown(source: str | Path) -> Path:
    """接受导出目录或 .md 文件。"""
    path = Path(source)
    if path.is_file():
        if path.suffix.lower() != ".md":
            raise SpecError(f"{path} 不是 .md 文件")
        return path
    if not path.is_dir():
        raise SpecError(f"找不到 {path}")
    candidates = sorted(path.glob("*.md"))
    if not candidates:
        raise SpecError(f"{path} 下没有 .md 文件")
    if len(candidates) > 1:
        preferred = path / f"{path.name}.md"
        if preferred.is_file():
            return preferred
    return candidates[0]


def default_output(markdown: Path) -> Path:
    return markdown.parent / f"{markdown.stem}.pptx"


def make_theme(accent: str | None = None) -> Theme:
    if not accent:
        return DEFAULT_THEME
    value = accent.strip().lstrip("#")
    if len(value) != 6:
        raise SpecError(f"--accent 需要 6 位十六进制颜色，收到 {accent!r}")
    try:
        rgb = RGBColor.from_string(value.upper())
    except Exception as exc:  # pragma: no cover
        raise SpecError(f"--accent 不是合法颜色：{accent!r}") from exc
    return replace(DEFAULT_THEME, accent=rgb)


def rebuild_from_export(
    source: str | Path,
    *,
    out: str | Path | None = None,
    settings: Settings | None = None,
    use_cache: bool = True,
    cache_dir: str | Path | None = None,
    concurrency: int | None = None,
    save_spec: str | Path | None = None,
    dump_dir: str | Path | None = None,
    accent: str | None = None,
    progress: bool = True,
) -> RebuildResult:
    """读取导出目录中的 Markdown，让模型生成版式规格，再渲染成 PPTX。"""
    markdown = locate_markdown(source)
    settings = settings or load_settings()
    if concurrency:
        settings = replace(settings, concurrency=concurrency)
    if cache_dir:
        settings = replace(settings, cache_dir=Path(cache_dir))

    if progress:
        print(
            f"[ppt-king] 还原 {markdown.name}，模型 {settings.model} @ {settings.base_url}",
            file=sys.stderr,
        )

    client = LLMClient(settings)
    cache = Cache.create(settings.cache_dir, enabled=use_cache)
    spec = compose_deck(
        markdown,
        settings=settings,
        client=client,
        cache=cache,
        concurrency=settings.concurrency,
        progress=progress,
        dump_dir=Path(dump_dir) if dump_dir else None,
    )

    spec_path = spec.save(save_spec) if save_spec else None
    output = Path(out) if out else default_output(markdown)
    builder.build_deck(spec, output, theme=make_theme(accent), base_dir=markdown.parent)

    if progress:
        print(
            f"[ppt-king] 已还原 {len(spec.slides)} 页 -> {output}\n[ppt-king] {client.stats()}",
            file=sys.stderr,
        )
    return RebuildResult(
        markdown=markdown,
        output=output,
        spec=spec,
        spec_path=spec_path,
        stats=client.stats(),
    )


def rebuild_from_spec(
    spec_path: str | Path,
    *,
    out: str | Path | None = None,
    base_dir: str | Path | None = None,
    accent: str | None = None,
    progress: bool = True,
) -> RebuildResult:
    """直接用已保存的 spec.json 渲染，不调用模型。"""
    path = Path(spec_path)
    spec = DeckSpec.load(path)
    root = Path(base_dir) if base_dir else path.parent
    output = Path(out) if out else root / f"{path.stem}.pptx"
    builder.build_deck(spec, output, theme=make_theme(accent), base_dir=root)
    if progress:
        print(f"[ppt-king] 已按 spec 渲染 {len(spec.slides)} 页 -> {output}", file=sys.stderr)
    return RebuildResult(markdown=None, output=output, spec=spec, spec_path=path)
