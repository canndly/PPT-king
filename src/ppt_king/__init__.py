"""ppt-king：从资料生成 PPT，并在 PPTX 与 Markdown 之间转换。"""

from __future__ import annotations

from .exporter import ExportResult, export_pptx
from .generate import GenerateResult, generate_deck
from .reader import read_pptx

__all__ = [
    "ExportResult",
    "GenerateResult",
    "export_pptx",
    "generate_deck",
    "read_pptx",
    "main",
]

__version__ = "0.2.0"


def main() -> None:
    from .cli import main as _main

    raise SystemExit(_main())
