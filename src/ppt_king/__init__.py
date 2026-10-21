"""ppt-king：PPTX <-> Markdown 双向转换（格式映射由大模型完成）。"""

from __future__ import annotations

from .exporter import ExportResult, export_pptx
from .reader import read_pptx
from .rebuild import RebuildResult, rebuild_from_export, rebuild_from_spec
from .spec import DeckSpec

__all__ = [
    "DeckSpec",
    "ExportResult",
    "RebuildResult",
    "export_pptx",
    "read_pptx",
    "rebuild_from_export",
    "rebuild_from_spec",
    "main",
]

__version__ = "0.1.0"


def main() -> None:
    from .cli import main as _main

    raise SystemExit(_main())
