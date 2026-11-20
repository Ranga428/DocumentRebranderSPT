"""Plain text, Markdown and HTML ingestion."""
from __future__ import annotations

import re
from datetime import date
from html.parser import HTMLParser

from .models import Block, DocumentModel, Run

_INLINE = re.compile(
    r"(\*\*.+?\*\*|__.+?__|\*[^*\n]+?\*|(?<!\w)_[A-Za-z0-9][^_\n]*?_(?!\w)|\[[^\]]+\]\([^)]+\))")


def inline_runs(text: str) -> list[Run]:
    runs: list[Run] = []
    for part in _INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            runs.append(Run(text=part[2:-2], bold=True))
        elif part.startswith("__") and part.endswith("__") and len(part) > 4:
            runs.append(Run(text=part[2:-2], underline=True))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            runs.append(Run(text=part[1:-1], italic=True))
        elif part.startswith("_") and part.endswith("_") and len(part) > 2:
            runs.append(Run(text=part[1:-1], italic=True))
        elif part.startswith("[") and "](" in part and part.endswith(")"):
            label, url = part[1:-1].split("](", 1)
            runs.append(Run(text=label, link=url))
        else:
            runs.append(Run(text=part))
    return runs or [Run(text=text)]


def _model_from_blocks(name: str, blocks: list[Block], warnings: list[str]) -> DocumentModel:
    model = DocumentModel(source_name=name, blocks=[b for b in blocks if b.text or b.kind in ("table", "hr", "image", "pagebreak")])
    heading = next((b for b in model.blocks if b.kind == "heading" and b.text.strip()), None)
    model.meta["title"] = heading.text.strip() if heading else name.rsplit(".", 1)[0]
    if heading is not None:
        model.blocks.remove(heading)
    model.meta["date"] = date.today().strftime("%d %B %Y")
    model.warnings = warnings
    return model


def _table_rows(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    rows = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        row = lines[i].strip().strip("|")
        cells = [c.strip() for c in row.split("|")]
        if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
            rows.append(cells)
        i += 1
    return rows, i


def ingest_markdown(data: bytes, name: str) -> DocumentModel:
    text = data.decode("utf-8", "replace")
    lines = text.splitlines()
    blocks: list[Block] = []
    warnings: list[str] = []
    para: list[str] = []
    in_code = False

    def flush() -> None:
        if para:
            blocks.append(Block(kind="paragraph", runs=inline_runs(" ".join(para).strip())))
            para.clear()

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("```"):
            flush()
            in_code = not in_code
            i += 1
            continue
        if in_code:
            para.append(line)
            i += 1
            continue
        if not stripped:
            flush()
            i += 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            flush()
            blocks.append(Block(kind="heading", level=len(m.group(1)),
                                runs=inline_runs(m.group(2).strip())))
            i += 1
            continue
        if stripped.startswith("|") and i + 1 < len(lines):
            rows, nxt = _table_rows(lines, i)
            if rows and len(rows[0]) > 1:
                flush()
                blocks.append(Block(kind="table", rows=rows))
                i = nxt
                continue
        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", stripped):
            flush()
            i += 1
            continue
        m = re.match(r"^(\s*)(?:[-*+]|\u2022|\u00b7|\u25cf)\s*(.*)$", line)
        if m:
            flush()
            blocks.append(Block(kind="bullet", level=min(3, len(m.group(1)) // 2),
                                runs=inline_runs(m.group(2).strip())))
            i += 1
            continue
        m = re.match(r"^(\s*)\d+[.)]\s+(.*)$", line)
        if m:
            flush()
            blocks.append(Block(kind="number", level=min(3, len(m.group(1)) // 2),
                                runs=inline_runs(m.group(2).strip())))
            i += 1
            continue
        if stripped.startswith("!["):
            warnings.append("Embedded Markdown images are referenced by URL and were skipped.")
            i += 1
            continue
        para.append(stripped)
        i += 1
    flush()
    return _model_from_blocks(name, blocks, warnings)


def _is_label_line(line: str) -> bool:
    return bool(re.match(r"^[A-Z][A-Za-z0-9 _()/-]{1,40}:\s*\S", line))


def ingest_text(data: bytes, name: str) -> DocumentModel:
    for enc in ("utf-8", "utf-16", "cp1252"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:  # pragma: no cover
        text = data.decode("utf-8", "replace")

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]

    blocks: list[Block] = []
    para: list[str] = []

    def flush() -> None:
        if not para:
            return
        chunk = " ".join(" ".join(para).split())
        words = len(chunk.split())
        looks_heading = (
            len(para) == 1 and words <= 12 and len(chunk) <= 90
            and not chunk.endswith((".", ",", ";", ":"))
            and (chunk.isupper() or chunk[:1].isupper())
        )
        if looks_heading:
            blocks.append(Block(kind="heading", level=1, runs=[Run(text=chunk)]))
        else:
            blocks.append(Block(kind="paragraph", runs=[Run(text=chunk)]))
        para.clear()

    for raw in lines:
        line = raw.strip()
        if not line:
            flush()
            continue
        bc = "\u2022\u00b7\u25cf\u25cb\u25a0\u25aa\u25ba\u2023\u2043\uf0b7\uf0a7\uf076\uf0d8"
        if para and (_is_label_line(line) or _is_label_line(para[-1])
                    or re.match(r"^\d+\.", line)
                    or any(line.startswith(c) for c in bc)
                    or any(para[-1].startswith(c) for c in bc)):
            flush()
        para.append(line)
    flush()
    return _model_from_blocks(name, blocks, [])


class _HTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.blocks: list[Block] = []
        self.warnings: list[str] = []
        self.buf: list[str] = []
        self.mode: str | None = None
        self.level = 0
        self.title = ""
        self._in_title = False
        self._skip = 0
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def flush(self, kind: str | None = None, level: int = 0) -> None:
        text = " ".join("".join(self.buf).split())
        self.buf.clear()
        if text:
            self.blocks.append(Block(kind=kind or "paragraph", level=level,
                                     runs=inline_runs(text)))

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head", "header", "footer"):
            self.flush()
            self._skip += 1
            if tag == "head":
                self._in_title = False
            return
        if tag == "title":
            self._in_title = True
        if self._skip:
            return
        if re.fullmatch(r"h[1-6]", tag):
            self.flush()
            self.mode, self.level = "heading", int(tag[1])
        elif tag == "p":
            self.flush()
            self.mode = "paragraph"
        elif tag == "li":
            self.flush()
            self.mode = "bullet"
        elif tag == "br":
            self.buf.append(" ")
        elif tag == "table":
            self.flush()
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
        elif tag == "img":
            self.warnings.append("HTML images are not downloaded; only the image is skipped.")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head", "header", "footer"):
            self._skip = max(0, self._skip - 1)
            return
        if tag == "title":
            self._in_title = False
        if self._skip:
            return
        if re.fullmatch(r"h[1-6]", tag) or tag in ("p", "li"):
            self.flush(self.mode, self.level)
            self.mode = None
        elif tag in ("td", "th") and self._cell is not None and self._row is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._row is not None and self._table is not None:
            if any(self._row):
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.blocks.append(Block(kind="table", rows=self._table))
            self._table = None

    def handle_data(self, data):
        if self._in_title:
            self.title += data
            return
        if self._skip:
            return
        if self._cell is not None:
            self._cell.append(data)
        else:
            self.buf.append(data)


def ingest_html(data: bytes, name: str) -> DocumentModel:
    parser = _HTML()
    parser.feed(data.decode("utf-8", "replace"))
    parser.close()
    parser.flush()
    model = _model_from_blocks(name, parser.blocks, parser.warnings)
    if parser.title.strip():
        model.meta["title"] = " ".join(parser.title.split())
    return model
