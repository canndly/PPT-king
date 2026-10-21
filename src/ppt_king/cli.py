"""命令行入口。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import ConfigurationError, load_settings
from .exporter import export_pptx
from .generate import GenerateError, generate_deck
from .grid import GridSyntaxError, describe, load_grid
from .grid_render import make_theme as make_grid_theme
from .grid_render import render_deck
from .llm import LLMError
from .outline import OutlineError
from .sources import SourceError


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt-king",
        description="从资料生成 PPT，或在 PPTX 与 Markdown 之间转换。",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    generate = sub.add_parser(
        "generate", help="资料目录 -> 大纲 -> 逐页网格 Markdown -> PPTX"
    )
    generate.add_argument("source", help="资料目录（.md/.txt/图片）")
    generate.add_argument("--outline", help="显式指定大纲文件；默认自动查找 outline.md/大纲.md")
    generate.add_argument("-o", "--out", help="输出目录（默认 out/<资料目录名>）")
    generate.add_argument(
        "--slides", default="8-20", help="页数，如 12 或 10-15（默认 8-20，仅约束 AI 大纲）"
    )
    generate.add_argument("--language", default="zh", help="正文语言（默认 zh）")
    generate.add_argument("--accent", help="主题强调色 / 表头色，如 2E5C8A")
    generate.add_argument("--logo", help="覆盖默认 Logo（默认 static/page_logo.png）")
    generate.add_argument("--no-logo", action="store_true", help="不显示默认 Logo")
    generate.add_argument("--header", help="右上角 8×1 区域的页眉文字")
    generate.add_argument("--resume", action="store_true", help="复用已有大纲和逐页结果")
    generate.add_argument("--force", action="store_true", help="忽略已有中间产物，全部重做")
    generate.add_argument("--strict", action="store_true", help="任意一页失败即整体失败")
    generate.add_argument("--no-cache", action="store_true", help="禁用响应缓存")
    generate.add_argument("--cache-dir", help="缓存目录（默认系统临时目录）")
    generate.add_argument("--debug-dump", help="把每页请求与响应保存到指定目录")
    generate.add_argument("--model", help="覆盖 PPT_KING_MODEL")
    generate.add_argument("--base-url", dest="base_url", help="覆盖 PPT_KING_BASE_URL")
    generate.add_argument("--api-key", dest="api_key", help="覆盖 PPT_KING_API_KEY")
    generate.add_argument("--env-file", help="指定 .env 路径")

    parse = sub.add_parser("parse", help="解析 PPTX 并导出 Markdown + 图片资源")
    parse.add_argument("pptx", nargs="+", help="一个或多个 .pptx 文件")
    parse.add_argument("-o", "--out", default="export", help="输出根目录（默认 export/）")
    parse.add_argument("--model", help="覆盖 PPT_KING_MODEL")
    parse.add_argument("--base-url", dest="base_url", help="覆盖 PPT_KING_BASE_URL")
    parse.add_argument("--api-key", dest="api_key", help="覆盖 PPT_KING_API_KEY")
    parse.add_argument("--concurrency", type=int, help="并发页数（默认 4）")
    parse.add_argument("--no-notes", action="store_true", help="忽略演讲者备注")
    parse.add_argument(
        "--drop-repeated-images",
        action="store_true",
        help="忽略在多数页面重复出现的图片（模板装饰/logo）",
    )
    parse.add_argument("--no-cache", action="store_true", help="禁用响应缓存")
    parse.add_argument("--cache-dir", help="缓存目录（默认系统临时目录）")
    parse.add_argument("--debug-dump", help="把每页请求与响应保存到指定目录")
    parse.add_argument("--stdout", action="store_true", help="只打印 Markdown，不写 .md 文件")
    parse.add_argument("--env-file", help="指定 .env 路径")

    render = sub.add_parser(
        "render", help="把 16×9 网格 Markdown 渲染成 PPTX（不调用模型）"
    )
    render.add_argument("md", help=".md 文件")
    render.add_argument("-o", "--out", help="输出 .pptx（默认与 .md 同名）")
    render.add_argument("--accent", help="主题强调色 / 表头色，如 2E5C8A")
    render.add_argument("--logo", help="覆盖默认 Logo（默认 static/page_logo.png）")
    render.add_argument("--no-logo", action="store_true", help="不显示默认 Logo")
    render.add_argument("--header", help="右上角 8×1 区域的页眉文字")
    render.add_argument("--check", action="store_true", help="只打印解析结果，不渲染")

    layout = sub.add_parser(
        "layout", help="让模型给普通 Markdown 加上 16×9 网格坐标"
    )
    layout.add_argument("md", help="普通 .md 文件")
    layout.add_argument("-o", "--out", help="输出网格 .md（默认 <名字>.grid.md）")
    layout.add_argument("--render", dest="render_after", action="store_true", help="产出后直接渲染 PPTX")
    layout.add_argument("--accent", help="主题强调色 / 表头色，如 2E5C8A")
    layout.add_argument("--logo", help="覆盖默认 Logo（默认 static/page_logo.png）")
    layout.add_argument("--no-logo", action="store_true", help="不显示默认 Logo")
    layout.add_argument("--header", help="右上角 8×1 区域的页眉文字")
    layout.add_argument("--no-cache", action="store_true", help="禁用响应缓存")
    layout.add_argument("--cache-dir", help="缓存目录（默认系统临时目录）")
    layout.add_argument("--model", help="覆盖 PPT_KING_MODEL")
    layout.add_argument("--base-url", dest="base_url", help="覆盖 PPT_KING_BASE_URL")
    layout.add_argument("--api-key", dest="api_key", help="覆盖 PPT_KING_API_KEY")
    layout.add_argument("--env-file", help="指定 .env 路径")
    return parser


def _force_utf8() -> None:
    """Windows 控制台默认 GBK，遇到 ⇒ / ³ 等字符会崩，这里统一切到 UTF-8。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass


def _parse_slide_range(value: str) -> tuple[int, int]:
    text = value.strip().lower().replace("页", "")
    if not text:
        return 8, 20
    if "-" in text:
        left, right = text.split("-", 1)
    elif "~" in text:
        left, right = text.split("~", 1)
    else:
        left = right = text
    try:
        minimum, maximum = int(left), int(right)
    except ValueError as exc:
        raise ValueError(f"--slides 需要形如 12 或 10-15，收到 {value!r}") from exc
    if minimum < 1 or maximum < minimum:
        raise ValueError(f"--slides 页数范围不合法：{value!r}")
    return minimum, maximum


def _overrides(args: argparse.Namespace) -> dict[str, object]:
    return {
        "model": args.model,
        "base_url": args.base_url,
        "api_key": args.api_key,
        "cache_dir": Path(args.cache_dir) if getattr(args, "cache_dir", None) else None,
    }


def main(argv: list[str] | None = None) -> int:
    _force_utf8()
    args = _build_parser().parse_args(argv)

    if args.command == "generate":
        return _run_generate(args)
    if args.command == "parse":
        return _run_parse(args)
    if args.command == "render":
        return _run_render(args)
    if args.command == "layout":
        return _run_layout(args)
    return 1


def _run_generate(args: argparse.Namespace) -> int:
    try:
        min_pages, max_pages = _parse_slide_range(args.slides)
        settings = load_settings(args.env_file, overrides=_overrides(args))
        result = generate_deck(
            args.source,
            out=args.out,
            outline=args.outline,
            settings=settings,
            resume=args.resume,
            force=args.force,
            min_pages=min_pages,
            max_pages=max_pages,
            language=args.language,
            accent=args.accent,
            logo=False if args.no_logo else args.logo,
            header=args.header,
            use_cache=not args.no_cache,
            dump_dir=args.debug_dump,
            strict=args.strict,
        )
    except ConfigurationError as exc:
        print(f"[ppt-king] 配置错误：{exc}", file=sys.stderr)
        return 2
    except (SourceError, OutlineError, GenerateError, LLMError, ValueError) as exc:
        print(f"[ppt-king] 生成失败：{exc}", file=sys.stderr)
        return 1

    print(
        f"[ppt-king] 输出目录：{result.output_dir}\n"
        f"[ppt-king] 大纲：{result.outline_path}\n"
        f"[ppt-king] 网格 Markdown：{result.grid_path}\n"
        f"[ppt-king] PPTX：{result.pptx_path}\n"
        f"[ppt-king] {result.pages} 页，复用 {result.cached_pages} 页；{result.stats}",
        file=sys.stderr,
    )
    if result.failures:
        print("[ppt-king] 以下页面使用了占位内容，可修复后用 --resume 重跑：", file=sys.stderr)
        for failure in result.failures:
            print(f"  - {failure}", file=sys.stderr)
        return 1
    return 0


def _run_parse(args: argparse.Namespace) -> int:
    try:
        settings = load_settings(args.env_file, overrides=_overrides(args))
    except ConfigurationError as exc:
        print(f"[ppt-king] 配置错误：{exc}", file=sys.stderr)
        return 2

    failures = 0
    for raw in args.pptx:
        path = Path(raw)
        if not path.is_file():
            print(f"[ppt-king] 找不到文件：{path}", file=sys.stderr)
            failures += 1
            continue
        try:
            result = export_pptx(
                path,
                out_root=args.out,
                settings=settings,
                include_notes=not args.no_notes,
                use_cache=not args.no_cache,
                dump_dir=args.debug_dump,
                write_file=not args.stdout,
                progress=not args.stdout,
                drop_repeated=args.drop_repeated_images,
            )
        except (LLMError, ValueError) as exc:
            print(f"[ppt-king] {path.name} 解析失败：{exc}", file=sys.stderr)
            failures += 1
            continue
        except Exception as exc:  # noqa: BLE001 - 保证批量处理不中断
            print(f"[ppt-king] {path.name} 未预期错误：{exc!r}", file=sys.stderr)
            failures += 1
            continue

        if args.stdout:
            print(result.markdown)
        else:
            print(
                f"[ppt-king] 输出目录：{result.output_dir}\n"
                f"[ppt-king] Markdown：{result.markdown_path}\n"
                f"[ppt-king] 图片/附件：{len(result.deck.images)} 张图片\n"
                f"[ppt-king] {result.stats}",
                file=sys.stderr,
            )

    return 1 if failures else 0


def _run_render(args: argparse.Namespace) -> int:
    try:
        deck = load_grid(args.md)
    except GridSyntaxError as exc:
        print(f"[ppt-king] 语法错误：{exc}", file=sys.stderr)
        return 1

    if args.check:
        print(describe(deck))
        return 0

    try:
        theme = make_grid_theme(args.accent)
    except ValueError as exc:
        print(f"[ppt-king] {exc}", file=sys.stderr)
        return 2

    out = args.out or str(Path(args.md).with_suffix(".pptx"))
    target, report = render_deck(
        deck,
        out,
        theme=theme,
        base_dir=Path(args.md).parent,
        logo=False if args.no_logo else args.logo,
        header=args.header,
    )
    print(
        f"[ppt-king] {report.slides} 页 / {report.blocks} 块（其中 {report.auto_blocks} 块自动排版）"
        f" -> {target}",
        file=sys.stderr,
    )
    for item in report.missing_images:
        print(f"[ppt-king] 图片不存在：{item}（已用占位文字代替）", file=sys.stderr)
    for item in report.overflows[:10]:
        print(f"[ppt-king] 溢出警告：{item}", file=sys.stderr)
    return 0


def _run_layout(args: argparse.Namespace) -> int:
    from .grid_ai import layout_markdown

    try:
        settings = load_settings(args.env_file, overrides=_overrides(args))
    except ConfigurationError as exc:
        print(f"[ppt-king] 配置错误：{exc}", file=sys.stderr)
        return 2

    source = Path(args.md)
    out = Path(args.out) if args.out else source.with_name(f"{source.stem}.grid.md")
    try:
        text = layout_markdown(source, settings=settings, use_cache=not args.no_cache)
    except (LLMError, GridSyntaxError) as exc:
        print(f"[ppt-king] 排版失败：{exc}", file=sys.stderr)
        return 1

    out.write_text(text, encoding="utf-8")
    print(f"[ppt-king] 网格 Markdown：{out}", file=sys.stderr)

    if args.render_after:
        try:
            theme = make_grid_theme(args.accent)
            deck = load_grid(out)
        except (GridSyntaxError, ValueError) as exc:
            print(f"[ppt-king] 渲染失败：{exc}", file=sys.stderr)
            return 1
        target, report = render_deck(
            deck,
            out.with_suffix(".pptx"),
            theme=theme,
            base_dir=out.parent,
            logo=False if args.no_logo else args.logo,
            header=args.header,
        )
        print(
            f"[ppt-king] {report.slides} 页 / {report.blocks} 块 -> {target}", file=sys.stderr
        )
        for item in report.missing_images:
            print(
                f"[ppt-king] 图片不存在：{item}（已用占位文字代替）",
                file=sys.stderr,
            )
        for item in report.overflows[:10]:
            print(f"[ppt-king] 溢出警告：{item}", file=sys.stderr)
    return 0
