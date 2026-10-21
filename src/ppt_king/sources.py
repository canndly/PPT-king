"""资料摄入：扫描资料目录，读取 Markdown/TXT，收集图片，发现大纲。

第一版只支持 ``.md`` / ``.markdown`` / ``.txt`` 与常见图片格式。
PDF、DOCX、PPTX 通过 :func:`register_extractor` 预留扩展入口，暂不内置解析器。
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

TEXT_SUFFIXES = {".md", ".markdown", ".txt"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg"}
OUTLINE_NAMES = ("outline.md", "大纲.md", "outline.txt", "大纲.txt")
IGNORED_DIRS = {".git", ".svn", ".hg", "__pycache__", "node_modules", ".venv", "venv"}

MAX_TEXT_BYTES = 2_000_000
MAX_IMAGES = 200
MAX_SKIPPED_LISTED = 20

_TITLE_RE = re.compile(r"^#\s+(.+)$", re.MULTILINE)
_MD_IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
_SAFE_NAME_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff_-]+")


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
def read_text_file(path: Path) -> str:
    """按 UTF-8 / GB18030 依次尝试读取文本。"""
    raw = path.read_bytes()[:MAX_TEXT_BYTES]
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


_read_plain_text = read_text_file


_EXTRACTORS: dict[str, Callable[[Path], str]] = {
    ".md": _read_plain_text,
    ".markdown": _read_plain_text,
    ".txt": _read_plain_text,
}


def register_extractor(suffix: str, reader: Callable[[Path], str]) -> None:
    """注册新的文本抽取器，供后续支持 PDF/DOCX/PPTX。"""
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
        if path.suffix.lower() not in TEXT_SUFFIXES:
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
            text = _read_plain_text(path)
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


def _copy_images(
    paths: list[Path], out_dir: Path, captions: dict[Path, str]
) -> list[SourceImage]:
    if not paths:
        return []
    assets = out_dir / "assets"
    assets.mkdir(parents=True, exist_ok=True)

    seen: dict[str, SourceImage] = {}
    images: list[SourceImage] = []
    for path in paths[:MAX_IMAGES]:
        try:
            digest = hashlib.sha1(path.read_bytes()).hexdigest()
        except OSError:
            continue
        if digest in seen:
            continue
        suffix = path.suffix.lower() or ".png"
        filename = f"src{len(images) + 1:02d}_{_safe_stem(path.stem)}{suffix}"
        target = assets / filename
        try:
            shutil.copyfile(path, target)
        except OSError:
            continue
        image = SourceImage(
            rel=f"assets/{filename}",
            path=target,
            title=path.stem,
            caption=captions.get(path.resolve(), ""),
        )
        seen[digest] = image
        images.append(image)
    return images


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

    outline_names = {name.lower() for name in OUTLINE_NAMES}
    for path in _iter_files(source, excludes):
        suffix = path.suffix.lower()
        if path.name.lower() in outline_names:
            continue
        if outline_resolved and path.resolve() == outline_resolved:
            continue
        if suffix in IMAGE_SUFFIXES:
            image_paths.append(path)
            continue
        reader = _EXTRACTORS.get(suffix)
        if reader is None:
            if len(skipped) < MAX_SKIPPED_LISTED:
                skipped.append(path.relative_to(source).as_posix())
            continue
        try:
            text = reader(path).strip()
        except OSError as exc:
            skipped.append(f"{path.name}（读取失败：{exc}）")
            continue
        if not text:
            continue
        match = _TITLE_RE.search(text)
        title = match.group(1).strip() if match else path.stem
        docs.append(
            SourceDoc(
                path=path,
                rel=path.relative_to(source).as_posix(),
                title=title,
                text=text,
            )
        )

    if not docs:
        raise SourceError(
            f"{source} 中没有可读取的 .md/.txt 资料（当前支持："
            + "、".join(sorted(TEXT_SUFFIXES))
            + "）"
        )

    images = _copy_images(image_paths, target_dir, _image_captions(source, excludes))
    return SourceBundle(
        root=source,
        out_dir=target_dir,
        docs=docs,
        images=images,
        outline=outline_resolved,
        skipped=skipped,
    )


def describe_skipped(bundle: SourceBundle) -> str:
    if not bundle.skipped:
        return ""
    listed = "、".join(bundle.skipped)
    return f"已跳过 {len(bundle.skipped)} 个暂不支持的文件：{listed}"
