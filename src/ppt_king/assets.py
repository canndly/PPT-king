"""资源落盘：图片、音视频、嵌入对象。

产物布局::

    export/<stem>/
    ├── <stem>.md
    ├── assets/slide01_img01.png
    └── media/slide03_file01.mp4

相同内容的二进制只写一次（按 SHA-1 去重），但每次引用都会分配新的 asset id，
便于 LLM 按页引用而不重复落盘。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

_SAFE_EXT = re.compile(r"[^a-z0-9]+")


def _safe_ext(ext: str | None, default: str = "bin") -> str:
    if not ext:
        return default
    cleaned = _SAFE_EXT.sub("", ext.lower().lstrip("."))
    return cleaned or default


@dataclass
class AssetSink:
    """把 shape 里的二进制写进输出目录并分配 id。"""

    root: Path
    images_dir: str = "assets"
    media_dir: str = "media"
    _image_seq: int = 0
    _file_seq: int = 0
    _by_digest: dict[str, str] = field(default_factory=dict)
    written: list[Path] = field(default_factory=list)

    def prepare(self) -> None:
        (self.root / self.images_dir).mkdir(parents=True, exist_ok=True)
        (self.root / self.media_dir).mkdir(parents=True, exist_ok=True)

    # -- 图片 ---------------------------------------------------------------
    def add_image(self, blob: bytes, ext: str | None, slide_no: int) -> tuple[str, str]:
        """返回 ``(asset_id, 相对路径)``。"""
        self._image_seq += 1
        asset_id = f"IMG{self._image_seq}"
        digest = hashlib.sha1(blob).hexdigest()

        cached = self._by_digest.get(digest)
        if cached:
            return asset_id, cached

        self.prepare()
        filename = f"slide{slide_no:02d}_img{self._image_seq:02d}.{_safe_ext(ext, 'png')}"
        rel_path = f"{self.images_dir}/{filename}"
        target = self.root / rel_path
        target.write_bytes(blob)
        self._by_digest[digest] = rel_path
        self.written.append(target)
        return asset_id, rel_path

    # -- 其他文件 -----------------------------------------------------------
    def add_file(
        self, blob: bytes, ext: str | None, slide_no: int, default_ext: str = "bin"
    ) -> tuple[str, str]:
        self._file_seq += 1
        asset_id = f"FILE{self._file_seq}"
        digest = hashlib.sha1(blob).hexdigest()

        cached = self._by_digest.get(digest)
        if cached:
            return asset_id, cached

        self.prepare()
        filename = f"slide{slide_no:02d}_file{self._file_seq:02d}.{_safe_ext(ext, default_ext)}"
        rel_path = f"{self.media_dir}/{filename}"
        target = self.root / rel_path
        target.write_bytes(blob)
        self._by_digest[digest] = rel_path
        self.written.append(target)
        return asset_id, rel_path
