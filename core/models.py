"""Neutral document model shared by every ingest adapter and the renderer."""
from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from typing import Any


def _bid() -> str:
    return uuid.uuid4().hex[:10]


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    link: str | None = None


@dataclass
class Block:
    kind: str = "paragraph"
    level: int = 0
    runs: list[Run] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    header_row: bool = True
    image: bytes | None = None
    width_px: int | None = None
    id: str = field(default_factory=_bid)
    # paragraph spacing (points), None = use template default
    space_before: float | None = None
    space_after: float | None = None
    line_spacing: float | None = None
    indent_left: float | None = None
    indent_right: float | None = None
    first_line: float | None = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def set_text(self, text: str) -> None:
        """Replace the content with one plain run, keeping the first run's emphasis."""
        first = self.runs[0] if self.runs else Run()
        self.runs = [Run(text=text, bold=first.bold, italic=first.italic,
                         underline=first.underline, link=first.link)]

    @property
    def plain(self) -> bool:
        return all(not (r.bold or r.italic or r.underline or r.link) for r in self.runs)


@dataclass
class DocumentModel:
    source_name: str = ""
    meta: dict[str, Any] = field(default_factory=dict)
    blocks: list[Block] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    # -- convenience -------------------------------------------------------
    @property
    def title(self) -> str:
        return str(self.meta.get("title") or "")

    @title.setter
    def title(self, value: str) -> None:
        self.meta["title"] = value

    @property
    def subtitle(self) -> str:
        return str(self.meta.get("subtitle") or "")

    @subtitle.setter
    def subtitle(self, value: str) -> None:
        self.meta["subtitle"] = value

    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for b in self.blocks:
            out[b.kind] = out.get(b.kind, 0) + 1
        return out

    def words(self) -> int:
        total = 0
        for b in self.blocks:
            if b.kind == "table":
                total += sum(len(str(c).split()) for row in b.rows for c in row)
            elif b.kind == "image":
                continue
            else:
                total += len(b.text.split())
        return total

    def headings(self) -> list[tuple[int, str]]:
        return [(b.level, b.text) for b in self.blocks if b.kind == "heading" and b.text.strip()]

    def has_wide_content(self) -> bool:
        for b in self.blocks:
            if b.kind == "table" and b.rows and len(b.rows[0]) >= 5:
                return True
            if b.kind == "image" and b.width_px and b.width_px > 900:
                return True
        return False

    def summary(self) -> str:
        return ", ".join(f"{v} {k}{'' if v == 1 else 's'}" for k, v in sorted(self.counts().items())) or "empty"
