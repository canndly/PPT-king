"""版式规格（DeckSpec）：大模型输出的 JSON 与渲染器之间的契约。

设计原则与解析方向一致：模型只决定**结构意图**（版式、块类型、文字、图片引用），
不接触任何真实文件路径；图片必须引用页内 `A1`/`A2` 这类 ref，由本地映射成文件。
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

LAYOUTS = {
    "title",
    "section",
    "title_content",
    "two_column",
    "image_left",
    "gallery",
    "table",
}

BLOCK_TYPES = {"heading", "paragraph", "bullets", "image", "table", "quote"}

_JSON_FENCE = re.compile(r"^```[a-zA-Z]*\s*|\s*```$")


class SpecError(ValueError):
    """模型输出的版式 JSON 不可用。"""


@dataclass
class BulletItem:
    text: str
    level: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {"text": self.text, "level": self.level}

    @classmethod
    def from_dict(cls, obj: Any) -> "BulletItem":
        if isinstance(obj, str):
            return cls(text=obj.strip())
        if not isinstance(obj, dict):
            raise SpecError(f"bullets 条目必须是对象或字符串，收到 {type(obj).__name__}")
        text = str(obj.get("text") or "").strip()
        try:
            level = int(obj.get("level") or 0)
        except (TypeError, ValueError):
            level = 0
        return cls(text=text, level=max(0, min(level, 4)))


@dataclass
class Block:
    type: str
    text: str = ""
    items: list[BulletItem] = field(default_factory=list)
    ref: str = ""
    caption: str = ""
    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"type": self.type}
        if self.text:
            data["text"] = self.text
        if self.items:
            data["items"] = [item.to_dict() for item in self.items]
        if self.ref:
            data["ref"] = self.ref
        if self.caption:
            data["caption"] = self.caption
        if self.header:
            data["header"] = list(self.header)
        if self.rows:
            data["rows"] = [list(row) for row in self.rows]
        return data

    @classmethod
    def from_dict(cls, obj: Any, allowed_refs: set[str] | None = None) -> "Block":
        if not isinstance(obj, dict):
            raise SpecError(f"block 必须是对象，收到 {type(obj).__name__}")

        block_type = str(obj.get("type") or "").strip().lower()
        if block_type not in BLOCK_TYPES:
            raise SpecError(f"未知的 block 类型：{block_type!r}，允许 {sorted(BLOCK_TYPES)}")

        if block_type == "bullets":
            raw_items = obj.get("items") or []
            if isinstance(raw_items, str):
                raw_items = [line for line in raw_items.splitlines() if line.strip()]
            items = [BulletItem.from_dict(item) for item in raw_items]
            items = [item for item in items if item.text]
            if not items:
                raise SpecError("bullets 块没有任何条目")
            return cls(type=block_type, items=items)

        if block_type == "image":
            ref = str(obj.get("ref") or "").strip().upper()
            if not ref:
                raise SpecError("image 块缺少 ref")
            if allowed_refs is not None and ref not in allowed_refs:
                raise SpecError(
                    f"image 块引用了不存在的图片 {ref}，本页可用：{sorted(allowed_refs) or '无'}"
                )
            return cls(type=block_type, ref=ref, caption=str(obj.get("caption") or "").strip())

        if block_type == "table":
            header = _as_str_list(obj.get("header"))
            rows = [_as_str_list(row) for row in (obj.get("rows") or []) if isinstance(row, list)]
            if not header:
                raise SpecError("table 块缺少表头 header")
            cols = max([len(header), *(len(row) for row in rows)])
            header = header + [""] * (cols - len(header))
            rows = [row + [""] * (cols - len(row)) for row in rows]
            return cls(type=block_type, header=header, rows=rows)

        text = str(obj.get("text") or "").strip()
        if not text:
            raise SpecError(f"{block_type} 块缺少 text")
        return cls(type=block_type, text=text)


@dataclass
class SlideSpec:
    number: int = 0
    layout: str = "title_content"
    title: str = ""
    subtitle: str = ""
    blocks: list[Block] = field(default_factory=list)
    notes: str = ""
    assets: dict[str, str] = field(default_factory=dict)  # ref -> 相对路径

    def to_dict(self) -> dict[str, Any]:
        return {
            "number": self.number,
            "layout": self.layout,
            "title": self.title,
            "subtitle": self.subtitle,
            "blocks": [block.to_dict() for block in self.blocks],
            "notes": self.notes,
            "assets": dict(self.assets),
        }

    @classmethod
    def from_dict(cls, obj: Any) -> "SlideSpec":
        if not isinstance(obj, dict):
            raise SpecError("slide 必须是对象")
        layout = str(obj.get("layout") or "").strip().lower()
        if layout not in LAYOUTS:
            layout = "title_content"
        try:
            number = int(obj.get("number") or 0)
        except (TypeError, ValueError):
            number = 0
        assets = {
            str(key).upper(): str(value)
            for key, value in (obj.get("assets") or {}).items()
            if isinstance(key, str) and isinstance(value, str)
        }
        blocks: list[Block] = []
        for raw in obj.get("blocks") or []:
            blocks.append(Block.from_dict(raw, allowed_refs=set(assets) or None))
        return cls(
            number=number,
            layout=layout,
            title=str(obj.get("title") or "").strip(),
            subtitle=str(obj.get("subtitle") or "").strip(),
            blocks=blocks,
            notes=str(obj.get("notes") or "").strip(),
            assets=assets,
        )


@dataclass
class DeckSpec:
    title: str = "未命名演示文稿"
    slides: list[SlideSpec] = field(default_factory=list)
    source: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": 1,
            "source": self.source,
            "title": self.title,
            "slides": [slide.to_dict() for slide in self.slides],
        }

    @classmethod
    def from_dict(cls, obj: Any) -> "DeckSpec":
        if not isinstance(obj, dict):
            raise SpecError("deck 必须是对象")
        slides = [SlideSpec.from_dict(item) for item in (obj.get("slides") or [])]
        if not slides:
            raise SpecError("deck 没有任何幻灯片")
        return cls(
            title=str(obj.get("title") or "未命名演示文稿").strip(),
            slides=slides,
            source=str(obj.get("source") or ""),
        )

    def save(self, path: str | Path) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(
            json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return target

    @classmethod
    def load(cls, path: str | Path) -> "DeckSpec":
        raw = Path(path).read_text(encoding="utf-8")
        try:
            return cls.from_dict(json.loads(raw))
        except json.JSONDecodeError as exc:
            raise SpecError(f"{path} 不是合法 JSON：{exc}") from exc


def _as_str_list(value: Any) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise SpecError("表格的 header/rows 必须是数组")
    return [str(item).strip() if item is not None else "" for item in value]


def extract_json(text: str) -> dict[str, Any]:
    """从模型输出里取出 JSON 对象（容忍代码围栏与前后废话）。"""
    cleaned = text.strip()
    cleaned = _JSON_FENCE.sub("", cleaned).strip()
    if not cleaned.startswith("{"):
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start == -1 or end <= start:
            raise SpecError("模型输出里找不到 JSON 对象")
        cleaned = cleaned[start : end + 1]
    try:
        obj = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise SpecError(f"模型输出的 JSON 无法解析：{exc}") from exc
    if not isinstance(obj, dict):
        raise SpecError("模型输出的 JSON 顶层必须是对象")
    return obj


def _normalize_for_compare(text: str) -> str:
    return re.sub(r"[\s，,、。．.:：；;！!？?\-—_]+", "", text).lower()


def _drop_duplicate_title(slide: SlideSpec) -> None:
    """模型常把页标题再写成正文，这里做确定性去重（跳过图片块）。"""
    target = _normalize_for_compare(slide.title)
    if not target:
        return
    for index, block in enumerate(slide.blocks):
        if block.type in ("image", "table", "quote"):
            continue
        if block.type in ("paragraph", "heading"):
            if _normalize_for_compare(block.text) == target:
                slide.blocks.pop(index)
            return
        if block.type == "bullets":
            if block.items and _normalize_for_compare(block.items[0].text) == target:
                block.items.pop(0)
                if not block.items:
                    slide.blocks.pop(index)
            return


def parse_slide(
    raw_text: str,
    *,
    number: int,
    allowed_refs: set[str],
    assets: dict[str, str],
    fallback_title: str = "",
) -> SlideSpec:
    """把模型输出解析为 SlideSpec，任何结构性错误都抛 SpecError（用于触发重问）。"""
    obj = extract_json(raw_text)
    slide = SlideSpec.from_dict({**obj, "number": number, "assets": assets})
    for block in slide.blocks:
        if block.type == "image" and block.ref not in allowed_refs:
            raise SpecError(f"image 块引用了本页不存在的图片 {block.ref}")
    if not slide.title:
        slide.title = fallback_title
    _drop_duplicate_title(slide)
    return slide
