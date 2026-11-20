"""Ingestion dispatcher: any supported file -> DocumentModel."""
from __future__ import annotations

import csv
import io
import re
from pathlib import Path

from . import office
from .ingest_docx import ingest_docx
from .ingest_pdf import ingest_pdf
from .ingest_text import ingest_html, ingest_markdown, ingest_text
from .models import Block, DocumentModel, Run



BULLET_CHARS = "\u2022\u00b7\u25cf\u25cb\u25a0\u25aa\u25ba\u2023\u2043\uf0b7\uf0a7\uf076\uf0d8"


def _split_inline_bullets(blocks):
    out = []
    for b in blocks:
        if b.kind != "paragraph":
            out.append(b)
            continue
        text = b.text
        if not any(c in text for c in BULLET_CHARS):
            out.append(b)
            continue
        bc = next((c for c in BULLET_CHARS if c in text), None)
        if bc is None:
            out.append(b)
            continue
        if text.lstrip().startswith(bc):
            nb = Block(kind="bullet", level=b.level, runs=[Run(text=text.lstrip(bc + " ").strip())])
            nb.space_before = b.space_before
            nb.space_after = b.space_after
            out.append(nb)
            continue
        lead = text.split(bc, 1)[0].strip()
        rest = [pt.strip() for pt in text.split(bc)[1:] if pt.strip()]
        if lead:
            pb = Block(kind="paragraph", runs=[Run(text=lead)])
            pb.space_before = b.space_before
            pb.space_after = b.space_after
            out.append(pb)
        for pt in rest:
            bb = Block(kind="bullet", level=b.level, runs=[Run(text=pt)])
            bb.space_before = b.space_before
            bb.space_after = b.space_after
            out.append(bb)
    return out





# A numbered heading is "4. Title" or "4.1 Title". A bare quantity
# such as "10 MB limit" or "12.5 mm" is deliberately not a heading.
_NUMBER_WORD = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,3})(\.?)\s+([A-Z].*)$")

# Numbered heading that was glued mid-paragraph, always after sentence
# punctuation ("... before the apps. 4. Core licensing module").
_MID_HEADING = re.compile(
    r"(?<=[.!?])\s+(\d{1,2}(?:\.\d{1,2}){0,3}\.?\s+[A-Z][^\n.]{1,60}?)(?=\s+[A-Z]|$)")

_MERGE_PUNCT = (",", ":", ";", "-", "\u2013", "\u2014")

# "1. Use a proxy tool ..." is a procedure step, not a section heading.
_STEP_VERB = re.compile(
    r"^(?:use|navigate|open|go|click|run|send|enter|access|log|visit|"
    r"intercept|configure|enable|disable|set|ensure|identify|change|remove|"
    r"add|update|perform|check|review|create|delete|upload|download|install|"
    r"start|stop|select|browse|request|submit|replay|modify|observe|repeat|"
    r"restart|launch|inspect|capture|tap|swipe|login|log in|logout|"
    r"authenticate|register|attempt|exploit|bypass|test|verify|try)\b",
    re.I)


def _heading_number(text: str) -> str | None:
    """Return the numeric token when `text` opens like a section heading."""
    stripped = text.strip()
    m = _NUMBER_WORD.match(stripped)
    if not m:
        return None
    num, dot, rest = m.group(1), m.group(2), m.group(3)
    if not dot and "." not in num:
        return None
    if "." not in num and _STEP_VERB.match(rest):
        return None
    if len(stripped) > 80 or rest.rstrip().endswith((".", ",", ";", ":")):
        return None
    return num


def _is_numbered_heading(text: str) -> bool:
    return _heading_number(text) is not None


def _heading_segments(text: str) -> list[tuple[str, str]]:
    """Split text into ("text" | "heading", value) pieces at numbered headings."""
    segments: list[tuple[str, str]] = []
    for raw in str(text).split("\n"):
        line = raw.strip()
        if not line:
            continue
        if _is_numbered_heading(line):
            segments.append(("heading", line))
            continue
        rest = line
        while rest:
            m = _MID_HEADING.search(rest)
            if not m:
                segments.append(("text", rest))
                break
            before = rest[:m.start()].strip()
            if before:
                segments.append(("text", before))
            segments.append(("heading", m.group(1).strip()))
            rest = rest[m.end():].strip()
    return segments


def _split_numbered_headings(blocks):
    """Emit numbered headings that were glued into a paragraph or bullet."""
    out: list[Block] = []
    for b in blocks:
        if b.kind not in ("paragraph", "bullet") or not b.text.strip():
            out.append(b)
            continue
        segments = _heading_segments(b.text)
        if all(kind == "text" for kind, _ in segments):
            out.append(b)
            continue

        pending: list[str] = []
        text_kind = b.kind

        def flush():
            if not pending:
                return
            nb = Block(kind=text_kind, level=b.level, runs=[Run(text=" ".join(pending))])
            nb.space_before = b.space_before
            nb.space_after = b.space_after
            out.append(nb)
            pending.clear()

        for kind, value in segments:
            if kind == "heading":
                flush()
                num = re.match(r"^(\d+(?:\.\d+)*)", value).group(1)
                hb = Block(kind="heading", level=2 if "." not in num else 3,
                           runs=[Run(text=value, bold=True)])
                hb.space_before = b.space_before
                hb.space_after = b.space_after
                out.append(hb)
                text_kind = "paragraph"
            else:
                pending.append(value)
        flush()
    return out


def _continues_bullet(prev: Block, nxt: Block) -> bool:
    """True only when `nxt` really is a wrapped line of the bullet `prev`."""
    if nxt.kind != "paragraph" or not nxt.text.strip():
        return False
    text = nxt.text.strip()
    if _is_numbered_heading(text):
        return False
    # a bold line is its own heading / lead-in paragraph
    if nxt.runs and all(r.bold for r in nxt.runs if r.text.strip()):
        return False
    # a capitalised sentence after a finished sentence starts a new paragraph
    if text[0].isupper() and not prev.text.rstrip().endswith(_MERGE_PUNCT):
        return False
    return True


def _merge_bullet_continuations(blocks):
    out = []
    barrier = True
    for b in blocks:
        if b.kind == "gap":
            barrier = True
            continue
        if (not barrier and out and out[-1].kind == "bullet"
                and _continues_bullet(out[-1], b)):
            prev = out[-1]
            extra = b.text.strip().replace("\n", " ")
            if extra:
                prev.runs.append(Run(text=" " + extra))
            continue
        if b.kind == "bullet":
            for r in b.runs:
                r.text = r.text.replace("\n", " ").replace("  ", " ").strip()
        out.append(b)
        barrier = b.kind != "bullet"
    return out

MAX_UPLOAD = 100 * 1024 * 1024

SUPPORTED = {
    ".pdf": "PDF",
    ".docx": "Word",
    ".docm": "Word (macros ignored)",
    ".doc": "Word 97-2003",
    ".odt": "OpenDocument Text",
    ".rtf": "Rich Text",
    ".txt": "Plain text",
    ".md": "Markdown",
    ".markdown": "Markdown",
    ".html": "HTML",
    ".htm": "HTML",
    ".csv": "CSV table",
    ".xlsx": "Excel",
    ".xls": "Excel 97-2003",
    ".pptx": "PowerPoint",
    ".ppt": "PowerPoint 97-2003",
}

OFFICE_DOCX = {".doc", ".odt", ".rtf"}
OFFICE_PDF = {".xlsx", ".xls", ".pptx", ".ppt"}


class IngestError(RuntimeError):
    pass


def accept_types() -> list[str]:
    return [s.lstrip(".") for s in SUPPORTED]


def _csv_model(data: bytes, name: str) -> DocumentModel:
    text = data.decode("utf-8-sig", "replace")
    rows = [[c.strip() for c in row] for row in csv.reader(io.StringIO(text))]
    rows = [r for r in rows if any(r)]
    model = DocumentModel(source_name=name, meta={"title": name.rsplit(".", 1)[0]})
    if rows:
        model.blocks.append(Block(kind="table", rows=rows, header_row=True))
    else:
        model.warnings.append("The CSV file contained no rows.")
    return model


def ingest(filename: str, data: bytes) -> DocumentModel:
    if not data:
        raise IngestError("The file is empty.")
    if len(data) > MAX_UPLOAD:
        raise IngestError(
            f"The file is larger than the {MAX_UPLOAD // (1024 * 1024)} MB limit.")

    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED:
        raise IngestError(
            f"Unsupported file type '{suffix or '(none)'}'. Supported: "
            + ", ".join(sorted(SUPPORTED)))

    model: DocumentModel
    try:
        if suffix in (".docx", ".docm"):
            model = ingest_docx(data, filename)
            if suffix == ".docm":
                model.warnings.append("Macros in the .docm file were ignored.")
        elif suffix in OFFICE_DOCX:
            converted = office.to_docx(data, filename)
            model = ingest_docx(converted, filename)
            model.warnings.insert(0, f"Converted from {suffix.lstrip('.').upper()} with LibreOffice; "
                                     "check tables and spacing in the preview.")
        elif suffix == ".pdf":
            model = ingest_pdf(data, filename)
        elif suffix in (".md", ".markdown"):
            model = ingest_markdown(data, filename)
        elif suffix in (".html", ".htm"):
            model = ingest_html(data, filename)
        elif suffix == ".csv":
            model = _csv_model(data, filename)
        elif suffix in OFFICE_PDF:
            converted = office.to_pdf(data, filename)
            model = ingest_pdf(converted, filename)
            model.warnings.insert(0, f"{suffix.lstrip('.').upper()} files are converted to PDF first, "
                                     "then re-read; this is the lowest-fidelity path.")
        else:
            model = ingest_text(data, filename)
    except (office.OfficeUnavailable, office.ConversionError) as exc:
        raise IngestError(str(exc)) from exc
    except IngestError:
        raise
    except Exception as exc:
        raise IngestError(
            f"Could not read '{filename}'. The file may be damaged or use an "
            f"unsupported variant. ({exc})") from exc

    model.meta.setdefault("source_format", suffix.lstrip("."))
    model.source_name = filename
    model.meta.setdefault("has_cover", False)
    model.blocks = _split_inline_bullets(model.blocks)
    model.blocks = _split_numbered_headings(model.blocks)
    model.blocks = _merge_bullet_continuations(model.blocks)
    if not model.blocks:
        model.warnings.append("No content was extracted. Try a different file or "
                              "check that the document is not empty.")
    return model
