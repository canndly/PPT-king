"""模板库：发现、读取和组装模板推荐目录。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

DEFAULT_TEMPLATE_DIR = Path(__file__).resolve().parents[2] / "static" / "templates"
TEMPLATE_SUFFIXES = {".md", ".markdown"}


@dataclass(frozen=True)
class TemplateInfo:
    name: str
    path: Path
    content: str


def discover_templates(root: str | Path | None = None) -> list[Path]:
    """扫描模板目录，返回 Markdown 模板文件。"""
    base = Path(root) if root else DEFAULT_TEMPLATE_DIR
    if not base.is_dir():
        return []
    return sorted(
        path
        for path in base.rglob("*")
        if path.is_file() and path.suffix.lower() in TEMPLATE_SUFFIXES
    )


def load_templates(root: str | Path | None = None) -> list[TemplateInfo]:
    """读取模板内容；空文件或读失败的文件会被跳过。"""
    templates: list[TemplateInfo] = []
    for path in discover_templates(root):
        try:
            content = path.read_text(encoding="utf-8-sig").strip()
        except OSError:
            continue
        if content:
            templates.append(TemplateInfo(name=path.name, path=path, content=content))
    return templates


def catalog_text(templates: list[TemplateInfo]) -> str:
    """把模板内容拼成给模型看的目录。"""
    chunks: list[str] = []
    for template in templates:
        chunks.append(
            f"### {template.name}\n```markdown\n{template.content}\n```"
        )
    return "\n\n".join(chunks)
