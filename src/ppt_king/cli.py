"""命令行入口。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import ConfigurationError, load_settings
from .exporter import export_pptx
from .grid import GridSyntaxError, describe, load_grid
from .grid_render import make_theme as make_grid_theme
from .grid_render import render_deck
from .llm import LLMError
from .rebuild import rebuild_from_export, rebuild_from_spec
from .spec import SpecError


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt-king",
        description="把 PPTX 解析成 Markdown（格式映射由大模型完成）。",
    )
    sub = parser.add_subparsers(dest="command", required=True)

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

    build = sub.add_parser("build", help="用导出的 Markdown + 图片还原 PPTX（版式由模型决定）")
    build.add_argument("source", nargs="?", help="导出目录或 .md 文件（使用 --spec 时可省略）")
    build.add_argument("-o", "--out", help="输出 .pptx 路径（默认与 Markdown 同目录同名）")
    build.add_argument("--spec", help="从已保存的 spec.json 渲染，不调用模型")
    build.add_argument("--save-spec", help="把模型生成的版式 JSON 另存为文件")
    build.add_argument("--accent", help="强调色，如 1D4ED8")
    build.add_argument("--concurrency", type=int, help="并发页数（默认取配置）")
    build.add_argument("--no-cache", action="store_true", help="禁用响应缓存")
    build.add_argument("--cache-dir", help="缓存目录（默认系统临时目录）")
    build.add_argument("--debug-dump", help="把每页版式 JSON 保存到指定目录")
    build.add_argument("--model", help="覆盖 PPT_KING_MODEL")
    build.add_argument("--base-url", dest="base_url", help="覆盖 PPT_KING_BASE_URL")
    build.add_argument("--api-key", dest="api_key", help="覆盖 PPT_KING_API_KEY")
    build.add_argument("--env-file", help="指定 .env 路径")

    render = sub.add_parser(
        "render", help="把 16×9 网格 Markdown 渲染成 PPTX（不调用模型）"
    )
    render.add_argument("md", help=".md 文件")
    render.add_argument("-o", "--out", help="输出 .pptx（默认与 .md 同名）")
    render.add_argument("--accent", help="主题强调色 / 表头色，如 2E5C8A")
    render.add_argument("--check", action="store_true", help="只打印解析结果，不渲染")

    layout = sub.add_parser(
        "layout", help="让模型给普通 Markdown 加上 16×9 网格坐标"
    )
    layout.add_argument("md", help="普通 .md 文件")
    layout.add_argument("-o", "--out", help="输出网格 .md（默认 <名字>.grid.md）")
    layout.add_argument("--render", dest="render_after", action="store_true", help="产出后直接渲染 PPTX")
    layout.add_argument("--accent", help="主题强调色 / 表头色，如 2E5C8A")
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


def main(argv: list[str] | None = None) -> int:
    _force_utf8()
    args = _build_parser().parse_args(argv)

    if args.command == "parse":
        return _run_parse(args)
    if args.command == "build":
        return _run_build(args)
    if args.command == "render":
        return _run_render(args)
    if args.command == "layout":
        return _run_layout(args)
    return 1


def _run_parse(args: argparse.Namespace) -> int:
    overrides: dict[str, object] = {
        "model": args.model,
        "base_url": args.base_url,
        "api_key": args.api_key,
        "cache_dir": Path(args.cache_dir) if args.cache_dir else None,
    }
    try:
        settings = load_settings(args.env_file, overrides=overrides)
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


def _run_build(args: argparse.Namespace) -> int:
    if not args.spec and not args.source:
        print("[ppt-king] 需要提供导出目录/ .md 文件，或用 --spec 指定版式 JSON", file=sys.stderr)
        return 2

    try:
        if args.spec:
            result = rebuild_from_spec(
                args.spec,
                out=args.out,
                base_dir=Path(args.source) if args.source and Path(args.source).is_dir() else None,
                accent=args.accent,
            )
        else:
            overrides: dict[str, object] = {
                "model": args.model,
                "base_url": args.base_url,
                "api_key": args.api_key,
                "cache_dir": Path(args.cache_dir) if args.cache_dir else None,
            }
            settings = load_settings(args.env_file, overrides=overrides)
            result = rebuild_from_export(
                args.source,
                out=args.out,
                settings=settings,
                use_cache=not args.no_cache,
                concurrency=args.concurrency,
                save_spec=args.save_spec,
                dump_dir=args.debug_dump,
                accent=args.accent,
            )
    except ConfigurationError as exc:
        print(f"[ppt-king] 配置错误：{exc}", file=sys.stderr)
        return 2
    except (LLMError, SpecError, ValueError) as exc:
        print(f"[ppt-king] 还原失败：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"[ppt-king] 未预期错误：{exc!r}", file=sys.stderr)
        return 1

    print(f"[ppt-king] PPTX：{result.output}", file=sys.stderr)
    if result.spec_path:
        print(f"[ppt-king] 版式 JSON：{result.spec_path}", file=sys.stderr)
    return 0


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
    target, report = render_deck(deck, out, theme=theme, base_dir=Path(args.md).parent)
    print(
        f"[ppt-king] {report.slides} 页 / {report.blocks} 块（其中 {report.auto_blocks} 块自动排版）"
        f" -> {target}",
        file=sys.stderr,
    )
    for item in report.overflows[:10]:
        print(f"[ppt-king] 溢出警告：{item}", file=sys.stderr)
    return 0


def _run_layout(args: argparse.Namespace) -> int:
    from .grid_ai import layout_markdown

    overrides: dict[str, object] = {
        "model": args.model,
        "base_url": args.base_url,
        "api_key": args.api_key,
        "cache_dir": Path(args.cache_dir) if args.cache_dir else None,
    }
    try:
        settings = load_settings(args.env_file, overrides=overrides)
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
            deck, out.with_suffix(".pptx"), theme=theme, base_dir=out.parent
        )
        print(
            f"[ppt-king] {report.slides} 页 / {report.blocks} 块 -> {target}", file=sys.stderr
        )
        for item in report.overflows[:10]:
            print(f"[ppt-king] 溢出警告：{item}", file=sys.stderr)
    return 0
