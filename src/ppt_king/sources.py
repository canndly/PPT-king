"""资料摄入：扫描资料目录，读取多种格式资料，收集图片，发现大纲。

格式转换由 :mod:`ppt_king.documents` 负责，支持 txt/md/code/CSV/HTML/PDF/
DOCX/PPTX/XLSX/EPUB/ZIP/Notebook 等；文档内嵌图片会落盘到 ``assets/``。
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .documents import (
    IMAGE_SUFFIXES,
    PLAIN_TEXT_SUFFIXES,
    SUPPORTED_SUFFIXES,
    ConversionError,
    convert_document,
    read_text_file,
)

TEXT_SUFFIXES = set(SUPPORTED_SUFFIXES)
OUTLINE_NAMES = ("outline.md", "大纲.md", "outline.txt", "大纲.txt")
IGNORED_DIRS = {".git", ".svn", ".hg", "__pycache__", "node_modules", ".venv", "venv"}

MAX_IMAGES = 200
MAX_SKIPPED_LISTED = 20

_TITLE_RE = re.compile(r"^#\s+(.+)$", re.MULTILINE)
_MD_IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
_SAFE_NAME_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff_-]+")
_PLACEHOLDER_RE = re.compile(r"!\[([^\]]*)\]\(pptking-image://([^)\s]+)\)")


class SourceError(ValueError):
    """资料目录不可用或没有可读取的内容。"""


@dataclass(frozen=True)
class SourceImage:
    rel: str  # 相对输出目录，如 assets/src01_原理图.png
    path: Path  # 输出目录中的真实路径
    title: str
    caption: str = ""


@dataclass
class SourceDoc:
    path: Path
    rel: str
    title: str
    text: str


@dataclass
class SourceBundle:
    root: Path
    out_dir: Path
    docs: list[SourceDoc] = field(default_factory=list)
    images: list[SourceImage] = field(default_factory=list)
    outline: Path | None = None
    skipped: list[str] = field(default_factory=list)

    @property
    def image_map(self) -> dict[str, SourceImage]:
        return {image.rel: image for image in self.images}

    @property
    def text_chars(self) -> int:
        return sum(len(doc.text) for doc in self.docs)


# --------------------------------------------------------------------------
# 格式注册
# --------------------------------------------------------------------------
# 外部可以注册自定义抽取器；内置格式由 documents.convert_document 处理。
_EXTRACTORS: dict[str, Callable[[Path], str]] = {}


def register_extractor(suffix: str, reader: Callable[[Path], str]) -> None:
    """注册额外的纯文本抽取器；同名后缀会覆盖内置转换。"""
    key = suffix.lower()
    if not key.startswith("."):
        key = "." + key
    _EXTRACTORS[key] = reader
    TEXT_SUFFIXES.add(key)


def supported_suffixes() -> list[str]:
    return sorted(TEXT_SUFFIXES | IMAGE_SUFFIXES)


# --------------------------------------------------------------------------
# 扫描
# --------------------------------------------------------------------------
def _resolve_out_dir(root: Path, out_dir: Path | None) -> Path:
    if out_dir is None:
        return (Path("out") / root.name).resolve()
    return Path(out_dir).resolve()


def _iter_files(root: Path, exclude: set[Path]):
    for current, dirs, files in os.walk(root):
        current_path = Path(current)
        if _excluded(current_path, exclude):
            continue
        dirs[:] = sorted(
            name
            for name in dirs
            if not name.startswith(".") and name not in IGNORED_DIRS
        )
        dirs[:] = [
            name
            for name in dirs
            if (current_path / name).resolve() not in exclude
        ]
        for name in sorted(files):
            if name.startswith("."):
                continue
            path = current_path / name
            if path.resolve() in exclude:
                continue
            yield path


def _excluded(path: Path, excludes: set[Path]) -> bool:
    resolved = path.resolve()
    return any(
        resolved == item or resolved.is_relative_to(item) for item in excludes
    )


def discover_outline(
    root: Path,
    explicit: str | Path | None = None,
    *,
    exclude: set[Path] | None = None,
) -> Path | None:
    """优先使用显式大纲；否则在资料目录内按约定文件名查找。"""
    if explicit:
        path = Path(explicit)
        if not path.is_absolute():
            path = Path.cwd() / path
        if not path.is_file():
            raise SourceError(f"找不到大纲文件：{explicit}")
        if path.suffix.lower() not in PLAIN_TEXT_SUFFIXES:
            raise SourceError(f"大纲文件必须是纯文本：{path}")
        return path.resolve()

    wanted = {name.lower(): index for index, name in enumerate(OUTLINE_NAMES)}
    candidates: list[tuple[int, Path]] = []
    for path in root.iterdir():
        if not path.is_file():
            continue
        if exclude and _excluded(path, exclude):
            continue
        index = wanted.get(path.name.lower())
        if index is None:
            continue
        candidates.append((index, path))
    if not candidates:
        return None
    return min(candidates, key=lambda item: item[0])[1].resolve()


def _image_captions(root: Path, exclude: set[Path] | None = None) -> dict[Path, str]:
    """从 Markdown 的 ``![alt](path)`` 中提取图片说明，按绝对路径索引。"""
    captions: dict[Path, str] = {}
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in {".md", ".markdown"}:
            continue
        if exclude and _excluded(path, exclude):
            continue
        try:
            text = read_text_file(path)
        except OSError:
            continue
        for match in _MD_IMAGE_RE.finditer(text):
            alt = match.group(1).strip()
            if not alt:
                continue
            target = (path.parent / match.group(2)).resolve()
            captions.setdefault(target, alt)
    return captions


def _safe_stem(stem: str) -> str:
    cleaned = _SAFE_NAME_RE.sub("_", stem).strip("_")
    return cleaned[:40] or "image"


class _ImageSink:
    """把散落在资料中的图片统一写入 ``assets/`` 并按内容去重。"""

    def __init__(self, out_dir: Path) -> None:
        self.assets = out_dir / "assets"
        self.assets.mkdir(parents=True, exist_ok=True)
        self.images: list[SourceImage] = []
        self._seen: dict[str, SourceImage] = {}

    def _next_name(self, stem: str, suffix: str) -> str:
        return f"src{len(self.images) + 1:02d}_{_safe_stem(stem)}{suffix}"

    def add_bytes(
        self,
        data: bytes | None,
        suffix: str,
        *,
        stem: str = "image",
        title: str = "",
        caption: str = "",
    ) -> str | None:
        """写入图片并返回相对输出目录的路径；重复图片直接复用。"""
        if not data or len(self.images) >= MAX_IMAGES:
            return None
        digest = hashlib.sha1(data).hexdigest()
        if digest in self._seen:
            return self._seen[digest].rel
        suffix = (suffix or ".png").lower()
        if not suffix.startswith("."):
            suffix = "." + suffix
        filename = self._next_name(stem or title or "image", suffix)
        target = self.assets / filename
        try:
            target.write_bytes(data)
        except OSError:
            return None
        image = SourceImage(
            rel=f"assets/{filename}",
            path=target,
            title=title or stem,
            caption=caption,
        )
        self._seen[digest] = image
        self.images.append(image)
        return image.rel

    def add_file(self, path: Path, caption: str = "") -> str | None:
        try:
            data = path.read_bytes()
        except OSError:
            return None
        digest = hashlib.sha1(data).hexdigest()
        if digest in self._seen:
            return self._seen[digest].rel
        if len(self.images) >= MAX_IMAGES:
            return None
        suffix = path.suffix.lower() or ".png"
        filename = self._next_name(path.stem, suffix)
        target = self.assets / filename
        try:
            shutil.copyfile(path, target)
        except OSError:
            return None
        image = SourceImage(
            rel=f"assets/{filename}",
            path=target,
            title=path.stem,
            caption=caption,
        )
        self._seen[digest] = image
        self.images.append(image)
        return image.rel


def scan_sources(
    root: str | Path,
    *,
    out_dir: str | Path | None = None,
    outline: str | Path | None = None,
) -> SourceBundle:
    """扫描资料目录并复制图片到输出目录的 ``assets/``。"""
    source = Path(root).resolve()
    if not source.is_dir():
        raise SourceError(f"资料目录不存在：{root}")

    target_dir = _resolve_out_dir(source, Path(out_dir) if out_dir else None)
    if target_dir == source:
        raise SourceError("输出目录不能与资料目录相同")
    excludes = {target_dir}

    outline_path = discover_outline(source, outline, exclude=excludes)
    outline_resolved = outline_path.resolve() if outline_path else None

    docs: list[SourceDoc] = []
    image_paths: list[Path] = []
    skipped: list[str] = []
    sink = _ImageSink(target_dir)
    doc_index = 0

    def note_skip(message: str) -> None:
        if len(skipped) < MAX_SKIPPED_LISTED:
            skipped.append(message)

    outline_names = {name.lower() for name in OUTLINE_NAMES}
    for path in _iter_files(source, excludes):
        suffix = path.suffix.lower()
        rel = path.relative_to(source).as_posix()
        if path.name.lower() in outline_names:
            continue
        if outline_resolved and path.resolve() == outline_resolved:
            continue
        if suffix in IMAGE_SUFFIXES:
            image_paths.append(path)
            continue
        reader = _EXTRACTORS.get(suffix)
        if reader is not None:
            try:
                text = reader(path).strip()
            except OSError as exc:
                note_skip(f"{rel}（读取失败：{exc}）")
                continue
            converted_images = []
        elif suffix in TEXT_SUFFIXES:
            doc_index += 1
            try:
                converted = convert_document(path, key_prefix=f"s{doc_index:03d}_")
            except ConversionError as exc:
                note_skip(f"{rel}（{exc}）")
                continue
            except OSError as exc:
                note_skip(f"{rel}（读取失败：{exc}）")
                continue
            text = converted.markdown.strip()
            converted_images = converted.images
        else:
            note_skip(rel)
            continue
        if not text:
            continue
        match = _TITLE_RE.search(text)
        title = match.group(1).strip() if match else path.stem

        if converted_images:
            mapping: dict[str, str] = {}
            for image in converted_images:
                stored = sink.add_bytes(
                    image.data,
                    image.suffix,
                    stem=f"{path.stem}_{image.key}",
                    title=image.title or path.stem,
                    caption=image.alt,
                )
                if stored:
                    mapping[image.key] = stored
            text = _rewrite_placeholders(text, mapping)
        docs.append(
            SourceDoc(
                path=path,
                rel=rel,
                title=title,
                text=text,
            )
        )

    if not docs:
        raise SourceError(
            f"{source} 中没有可读取的资料（支持 md/txt/csv/html/pdf/docx/pptx/"
            "xlsx/epub/zip 等，详见 README“支持的资料格式”）"
        )

    captions = _image_captions(source, excludes)
    for path in image_paths:
        sink.add_file(path, captions.get(path.resolve(), ""))

    return SourceBundle(
        root=source,
        out_dir=target_dir,
        docs=docs,
        images=sink.images,
        outline=outline_resolved,
        skipped=skipped,
    )


def _rewrite_placeholders(text: str, mapping: dict[str, str]) -> str:
    """把转换层生成的内部图片占位符换成实际 ``assets/`` 路径。"""

    def replace(match: re.Match[str]) -> str:
        alt, key = match.group(1), match.group(2)
        rel = mapping.get(key)
        if rel:
            return f"![{alt}]({rel})"
        return alt or ""

    return _PLACEHOLDER_RE.sub(replace, text)


def describe_skipped(bundle: SourceBundle) -> str:
    if not bundle.skipped:
        return ""
    listed = "、".join(bundle.skipped)
    return f"已跳过 {len(bundle.skipped)} 个暂不支持的文件：{listed}"
