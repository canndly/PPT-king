"""Markdown 后处理：清洗、校验、占位符替换。"""

from __future__ import annotations

import re

from .ir import EmbedBlock, ImageBlock, Slide

PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z]+\d+)\s*\}\}")
RAW_PATH_RE = re.compile(r"\]\(\s*(?:\./)?(?:assets|media)/")
HEADING_RE = re.compile(r"^#{2,6}\s", re.MULTILINE)


def normalize(text: str) -> str:
    """去掉代码围栏、尾部页分隔线等模型常见多余输出。"""
    cleaned = text.strip()
    cleaned = re.sub(r"\n\s*-{3,}\s*$", "", cleaned).strip()
    if cleaned.startswith("```"):
        lines = cleaned.splitlines()[1:]
        while lines and not lines[-1].strip():
            lines.pop()
        if lines and lines[-1].strip().startswith("```"):
            lines.pop()
        cleaned = "\n".join(lines).strip()
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    cleaned = re.sub(r"(^#{2,6} .*)\n(?=\S)", r"\1\n\n", cleaned, flags=re.MULTILINE)
    return cleaned.strip()


def validate_page(page_md: str, allowed: set[str]) -> list[str]:
    """返回问题列表，空列表表示通过。"""
    problems: list[str] = []
    if not page_md.strip():
        problems.append("输出为空")
        return problems
    if "```" in page_md:
        problems.append("出现了代码围栏 ```，不允许")
    if RAW_PATH_RE.search(page_md):
        problems.append("直接写了 assets/ 或 media/ 路径，必须使用 {{IMGn}} / {{FILEn}} 占位符")
    if not HEADING_RE.search(page_md):
        problems.append("缺少页标题行 `## 页码. 标题`")
    for token in PLACEHOLDER_RE.findall(page_md):
        if token not in allowed:
            problems.append(f"引用了本页不存在的占位符 {{{{{token}}}}}")
    if "{{" in PLACEHOLDER_RE.sub("", page_md):
        problems.append("存在格式错误的占位符（应形如 {{IMG1}}）")
    return problems


def _image_markdown(block: ImageBlock) -> str:
    alt = (block.alt or "").replace("[", "(").replace("]", ")").strip()
    return f"![{alt}]({block.rel_path})"


def _embed_markdown(block: EmbedBlock) -> str:
    label = (block.label or block.asset_id).replace("[", "(").replace("]", ")")
    return f"[{label}]({block.rel_path})"


def resolve_placeholders(
    page_md: str,
    images: dict[str, ImageBlock],
    embeds: dict[str, EmbedBlock],
) -> str:
    def _replace(match: re.Match[str]) -> str:
        token = match.group(1)
        if token in images:
            return _image_markdown(images[token])
        if token in embeds:
            return _embed_markdown(embeds[token])
        return match.group(0)

    return PLACEHOLDER_RE.sub(_replace, page_md)


def missing_asset_ids(page_md: str, slide: Slide) -> list[str]:
    """模型漏掉的图片/附件，用于兜底补全。"""
    used = set(PLACEHOLDER_RE.findall(page_md))
    missing: list[str] = []
    for block in slide.blocks:
        if isinstance(block, (ImageBlock, EmbedBlock)) and block.asset_id not in used:
            missing.append(block.asset_id)
    return missing


def append_missing(page_md: str, missing: list[str]) -> str:
    if not missing:
        return page_md
    extra = "\n\n".join("{{" + asset_id + "}}" for asset_id in missing)
    return f"{page_md.rstrip()}\n\n{extra}\n"


def ensure_heading(page_md: str, number: int, title: str) -> str:
    if HEADING_RE.search(page_md):
        return page_md
    heading = f"## {number}. {title}" if title else f"## 幻灯片 {number}"
    return f"{heading}\n\n{page_md.lstrip()}"


def assemble(deck_title: str, pages: list[str]) -> str:
    body = "\n\n---\n\n".join(page.strip() for page in pages)
    return f"# {deck_title}\n\n{body}\n"
