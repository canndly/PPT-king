"""中间表示（IR）：PPTX 抽取结果，与 Markdown 渲染解耦。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    level: int = 0
    bullet: str | None = None  # "char" | "number" | None（None = 未显式设置）

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)

    @property
    def is_empty(self) -> bool:
        return not self.text.strip()


@dataclass
class ShapeBase:
    name: str = ""
    top: int = 0
    left: int = 0
    width: int = 0
    height: int = 0
    shape_kind: str = "shape"


@dataclass
class TextBlock(ShapeBase):
    paragraphs: list[Paragraph] = field(default_factory=list)
    role: str = "text"  # title | subtitle | body | textbox
    placeholder: str | None = None
    shape_kind: str = "text"


@dataclass
class TableBlock(ShapeBase):
    rows: list[list[str]] = field(default_factory=list)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)  # (row, col, rowspan, colspan)
    spanned: set[tuple[int, int]] = field(default_factory=set)
    shape_kind: str = "table"


@dataclass
class ImageBlock(ShapeBase):
    asset_id: str = ""  # IMG1
    rel_path: str = ""  # assets/slide01_img01.png
    alt: str = ""
    shape_kind: str = "image"

    @property
    def markdown_token(self) -> str:
        return "{{" + self.asset_id + "}}"


@dataclass
class ChartBlock(ShapeBase):
    chart_type: str = ""
    title: str = ""
    categories: list[str] = field(default_factory=list)
    series: list[tuple[str, list[str]]] = field(default_factory=list)
    shape_kind: str = "chart"


@dataclass
class EmbedBlock(ShapeBase):
    asset_id: str = ""
    rel_path: str = ""
    kind: str = "file"  # audio | video | ole | file
    label: str = ""
    shape_kind: str = "embed"


@dataclass
class Slide:
    number: int
    title: str = ""
    blocks: list[Any] = field(default_factory=list)
    notes: list[Paragraph] = field(default_factory=list)

    @property
    def images(self) -> list[ImageBlock]:
        return [block for block in self.blocks if isinstance(block, ImageBlock)]


@dataclass
class Deck:
    source: Path
    title: str = ""
    slide_width: int = 0
    slide_height: int = 0
    slides: list[Slide] = field(default_factory=list)

    @property
    def images(self) -> list[ImageBlock]:
        return [image for slide in self.slides for image in slide.images]
