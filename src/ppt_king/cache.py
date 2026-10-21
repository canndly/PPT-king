"""页级响应缓存，避免开发期重复计费。"""

from __future__ import annotations

import hashlib
import tempfile
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CACHE_DIR = Path(tempfile.gettempdir()) / "ppt-king-cache"


@dataclass
class Cache:
    directory: Path
    enabled: bool = True

    @classmethod
    def create(cls, directory: Path | None = None, enabled: bool = True) -> "Cache":
        path = Path(directory) if directory else DEFAULT_CACHE_DIR
        if enabled:
            path.mkdir(parents=True, exist_ok=True)
        return cls(directory=path, enabled=enabled)

    @staticmethod
    def key(*parts: str) -> str:
        digest = hashlib.sha256()
        for part in parts:
            digest.update(part.encode("utf-8"))
            digest.update(b"\x00")
        return digest.hexdigest()

    def get(self, key: str) -> str | None:
        if not self.enabled:
            return None
        target = self.directory / f"{key}.md"
        if target.is_file():
            try:
                return target.read_text(encoding="utf-8")
            except OSError:
                return None
        return None

    def put(self, key: str, content: str) -> None:
        if not self.enabled:
            return
        try:
            self.directory.mkdir(parents=True, exist_ok=True)
            (self.directory / f"{key}.md").write_text(content, encoding="utf-8")
        except OSError:
            pass
