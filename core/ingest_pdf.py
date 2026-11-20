"""PDF ingestion with font-size heading heuristics.

PDF is the least faithful source: it has no semantic structure, so headings,
reading order and tables are all inferred. Every guess that matters is reported
in ``model.warnings``.
"""
from __future__ import annotations

import re
from collections import Counter

import pymupdf as fitz

from .models import Block, DocumentModel

_DOTS = re.compile(r"^(.*?\S)\s*[.·\u2026]{2,}\s*(\d{1,4})\s*$")

BULLET_CHARS = "\u2022\u00b7\u25cf\u25cb\u25a0\u25aa\u25ba\u2023\u2043\uf0b7\uf0a7\uf076\uf0d8"
BULLET_RE = re.compile(r"^[" + BULLET_CHARS + r"\-\*\+]\s*")
O_BULLET_RE = re.compile(r"^[oO]\s+")


def _line_specs(page):
    data = page.get_text("dict")
    out = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [s for s in line.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            runs = []
            for sp in spans:
                font = sp.get("font", "")
                is_bold = bool(sp.get("flags", 0) & 16) or "bold" in font.lower() or "bd" in font.lower().split("-")[-1:]
                is_italic = bool(sp.get("flags", 0) & 2) or "italic" in font.lower() or "it" in font.lower().split("-")[-1:]
                runs.append({
                    "text": sp["text"],
                    "bold": is_bold,
                    "italic": is_italic,
                    "size": round(sp.get("size", 0), 1),
                })
            text = "".join(r["text"] for r in runs).strip()
            if not text:
                continue
            size = max(r["size"] for r in runs)
            bold = sum(1 for r in runs if r["bold"]) / len(runs) > 0.5
            italic = sum(1 for r in runs if r["italic"]) / len(runs) > 0.5
            out.append({
                "text": text,
                "bbox": line["bbox"],
                "size": round(size, 1),
                "bold": bold,
                "italic": italic,
                "runs": runs,
            })
    return out


def _inside(bbox, outer) -> bool:
    x0, y0, x1, y1 = bbox
    ox0, oy0, ox1, oy1 = outer
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return ox0 <= cx <= ox1 and oy0 <= cy <= oy1


def _cell_text_from_spans(page_lines, cell_bbox):
    cx0, cy0, cx1, cy1 = cell_bbox
    spans = []
    for ln in page_lines:
        for sp in ln.get("spans", []):
            if not sp.get("text", "").strip():
                continue
            sx0, sy0, sx1, sy1 = sp["bbox"]
            cx, cy = (sx0 + sx1) / 2, (sy0 + sy1) / 2
            if not (cx0 - 3 <= cx <= cx1 + 3 and cy0 - 3 <= cy <= cy1 + 3):
                continue
            spans.append((sx0, sy0, sx1, sy1, sp["text"]))
    if not spans:
        return ""
    spans.sort(key=lambda s: (s[1], s[0]))
    line_groups = []
    for sp in spans:
        if line_groups and abs(sp[1] - line_groups[-1][-1][1]) <= 4:
            line_groups[-1].append(sp)
        else:
            line_groups.append([sp])
    out_lines = []
    for grp in line_groups:
        grp.sort(key=lambda s: s[0])
        parts = []
        prev_x1 = None
        for sp in grp:
            if prev_x1 is not None and sp[0] - prev_x1 > 2:
                parts.append(" ")
            parts.append(sp[4])
            prev_x1 = sp[2]
        out_lines.append("".join(parts).strip())
    return " ".join(l for l in out_lines if l).strip()


def _extract_tables(page):
    try:
        finder = page.find_tables()
    except Exception:
        return []
    out = []
    line_list = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for ln in block.get("lines", []):
            line_list.append({"spans": ln.get("spans", []), "bbox": ln["bbox"]})
    for table in getattr(finder, "tables", []):
        try:
            rows = []
            cells = getattr(table, "cells", None)
            if cells and getattr(table, "row_count", 0) and getattr(table, "col_count", 0):
                xs = sorted({cb[0] for cb in cells})
                ys = sorted({cb[1] for cb in cells})
                grid = [["" for _ in range(table.col_count)] for _ in range(table.row_count)]
                for cb in cells:
                    ci = xs.index(cb[0]) if cb[0] in xs else 0
                    ri = ys.index(cb[1]) if cb[1] in ys else 0
                    if 0 <= ri < table.row_count and 0 <= ci < table.col_count:
                        grid[ri][ci] = _cell_text_from_spans(line_list, cb)
                rows = grid
            else:
                rows = table.extract()
                rows = [[(cell or "").replace("\n", " ").strip() for cell in row] for row in rows]
        except Exception:
            try:
                rows = table.extract()
                rows = [[(cell or "").replace("\n", " ").strip() for cell in row] for row in rows]
            except Exception:
                continue
        rows = [r for r in rows if any(r)]
        if len(rows) >= 2 and max(len(r) for r in rows) >= 2:
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            out.append((table.bbox, rows))
    return out

def _extract_images(doc, page):
    out = []
    page_area = abs(page.rect.width * page.rect.height)
    for info in page.get_image_info(xrefs=True):
        xref = info.get("xref", 0)
        bbox = info.get("bbox")
        if not xref or not bbox:
            continue
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        if w < 80 or h < 40:
            continue
        if (w * h) > page_area * 0.85:
            continue
        try:
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            blob = pix.tobytes("png")
            width_px = pix.width
            pix = None
        except Exception:
            continue
        out.append((bbox[1], Block(kind="image", image=blob, width_px=width_px)))
    return out


def _looks_separate(text: str) -> bool:
    t = text.strip()
    if re.match(r"^\d+\.", t):
        return True
    if (t.startswith("\u2022") or t.startswith("\u00b7")
        or t.startswith("- ") or t.startswith("* ")):
        return True
    if O_BULLET_RE.match(t):
        return True
    if re.match(r"^[A-Z][^.!?]{0,60}:\s*$", t) or re.match(r"^[A-Z][^.!?]{0,60}:\s+\S", t):
        return True
    return False


def _merge_lines(lines: list[dict]) -> list[dict]:
    lines = sorted(lines, key=lambda l: (round(l["bbox"][1], 1), l["bbox"][0]))
    out: list[dict] = []
    for line in lines:
        if out:
            prev = out[-1]
            gap = line["bbox"][1] - prev["bbox"][3]
            height = max(1.0, prev["bbox"][3] - prev["bbox"][1])
            same_size = abs(line["size"] - prev["size"]) <= 1.0
            prev_separate = _looks_separate(prev["text"])
            next_separate = _looks_separate(line["text"])
            bold_diff = prev.get("bold", False) != line.get("bold", False)
            if same_size and gap < height * 0.9 and not prev_separate and not next_separate and not bold_diff:
                if prev["text"].endswith("-") and line["text"][:1].islower():
                    prev["text"] += line["text"]
                else:
                    prev["text"] += "\n" + line["text"]
                prev["runs"] = prev.get("runs", []) + line.get("runs", [])
                prev["bbox"] = (prev["bbox"][0], prev["bbox"][1],
                                max(prev["bbox"][2], line["bbox"][2]),
                                max(prev["bbox"][3], line["bbox"][3]))
                prev["bold"] = (sum(1 for r in prev["runs"] if r["bold"]) / max(1, len(prev["runs"])) > 0.5)
                prev["italic"] = (sum(1 for r in prev["runs"] if r["italic"]) / max(1, len(prev["runs"])) > 0.5)
                continue
        newl = dict(line)
        newl.setdefault("runs", list(line.get("runs", [])))
        out.append(newl)
    return out


_NUMBER_WORD = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,3})(\.?)\s+([A-Z].*)$")

# Numbered procedure steps ("1. Use a proxy tool", "2. Navigate to ...") are
# list items, not section headings.
_STEP_VERB = re.compile(
    r"^(?:use|navigate|open|go|click|run|send|enter|access|log|visit|"
    r"intercept|configure|enable|disable|set|ensure|identify|change|remove|"
    r"add|update|perform|check|review|create|delete|upload|download|install|"
    r"start|stop|select|browse|request|submit|replay|modify|observe|repeat|"
    r"restart|launch|inspect|capture|tap|swipe|login|log in|logout|"
    r"authenticate|register|attempt|exploit|bypass|test|verify|try)\b",
    re.I)


def _heading_level(size: float, bold: bool, body: float, text: str) -> int:
    t = text.strip()
    m = _NUMBER_WORD.match(t)
    if m:
        num, dot, rest = m.group(1), m.group(2), m.group(3)
        numbered = bool(dot) or "." in num
        short = len(t) <= 80 and not rest.rstrip().endswith((".", ",", ";", ":"))
        if numbered and short:
            # "1. Use a proxy tool" is a procedure step, never a heading
            if "." not in num and _STEP_VERB.match(rest):
                return 0
            # bold or visually larger numbered lines are section headings
            if bold or size >= body * 1.05:
                return 2 if "." not in num else 3
            # "4.1 Controllers" non-bold, but only when it does not read as a step
            if "." in num and not _STEP_VERB.match(rest):
                return 3
    if size >= body * 1.55:
        return 1
    if size >= body * 1.28:
        return 2
    if size >= body * 1.12:
        return 3
    if bold and size > body * 1.02 and len(text) < 90 and not text.rstrip().endswith((".", ",", ";")):
        return 3
    return 0


def ingest_pdf(data: bytes, name: str) -> DocumentModel:
    doc = fitz.open(stream=data, filetype="pdf")
    model = DocumentModel(source_name=name)
    warnings: list[str] = []

    if doc.needs_pass:
        raise ValueError("The PDF is password protected. Remove the password and retry.")

    # body font size = most common span size across the document
    sizes: Counter[float] = Counter()
    page_lines = []
    for page in doc:
        lines = _line_specs(page)
        page_lines.append(lines)
        for line in lines:
            sizes[line["size"]] += len(line["text"])
    body = sizes.most_common(1)[0][0] if sizes else 11.0

    # repeated running headers/footers on the page edges
    edge_counts: Counter[str] = Counter()
    for page, lines in zip(doc, page_lines):
        height = page.rect.height or 1
        for line in lines:
            y = (line["bbox"][1] + line["bbox"][3]) / 2
            if y < height * 0.09 or y > height * 0.91:
                edge_counts[line["text"]] += 1
    boiler = {t for t, c in edge_counts.items() if c >= max(2, len(doc) * 0.4)}

    # also treat any text sitting in the page margins as header/footer chrome,
    # even when it is not repeated (e.g. single-page documents)
    def _is_edge(line, page) -> bool:
        height = page.rect.height or 1
        y = (line["bbox"][1] + line["bbox"][3]) / 2
        if y < height * 0.09 or y > height * 0.91:
            return True
        text = line["text"].lower()
        if y < height * 0.15 or y > height * 0.85:
            if (re.search(r"@[\w.]+", text) or re.search(r"\+?\d[\d\s\-()]{7,}", text)
                    or "tel" in text or "contact" in text or "email" in text
                    or "www." in text or "http" in text):
                return True
        return False

    blocks: list[Block] = []
    scanned_pages: list[int] = []
    multi_column_pages: list[int] = []

    for page_number, (page, lines) in enumerate(zip(doc, page_lines), start=1):
        if len("".join(l["text"] for l in lines)) < 20:
            scanned_pages.append(page_number)

        tables = _extract_tables(page)
        items: list[tuple[float, float, Block]] = []
        for bbox, rows in tables:
            items.append((bbox[1], bbox[0], Block(kind="table", rows=rows)))


        usable = [l for l in lines
                  if l["text"] not in boiler
                  and not _is_edge(l, page)
                  and not any(_inside(l["bbox"], t[0]) for t in tables)]
        body_lines = []
        for line in usable:
            text = line["text"]
            bullet_char = None
            for bc in BULLET_CHARS:
                if bc in text:
                    bullet_char = bc
                    break
            if bullet_char and not text.lstrip().startswith(bullet_char):
                sep = bullet_char
                lead = text.split(sep, 1)[0].strip()
                rest = [pt.strip() for pt in text.split(sep)[1:] if pt.strip()]
                if lead:
                    body_lines.append({
                        "text": lead,
                        "bbox": line["bbox"],
                        "size": line["size"],
                        "bold": line["bold"],
                        "italic": line["italic"],
                        "runs": [{"text": lead, "bold": line["bold"],
                                 "italic": line["italic"], "size": line["size"]}],
                    })
                for ptxt in rest:
                    body_lines.append({
                        "text": "\u2022 " + ptxt,
                        "bbox": line["bbox"],
                        "size": line["size"],
                        "bold": line["bold"],
                        "italic": line["italic"],
                        "runs": [{"text": "\u2022 " + ptxt, "bold": line["bold"],
                                 "italic": line["italic"], "size": line["size"]}],
                    })
                continue
            body_lines.append(line)

        for merged in _merge_lines(body_lines):
            text = merged["text"].strip()
            level = _heading_level(merged["size"], merged["bold"], body, text)
            bullet = False
            if BULLET_RE.match(text) or O_BULLET_RE.match(text) or (text and text[0] in BULLET_CHARS):
                bullet = True
                if BULLET_RE.match(text):
                    text = BULLET_RE.sub("", text).strip()
                elif O_BULLET_RE.match(text):
                    text = O_BULLET_RE.sub("", text).strip()
                else:
                    text = text.lstrip(BULLET_CHARS + "-*+").strip()
            runs = merged.get("runs") or [{"text": text, "bold": merged["bold"],
                                            "italic": merged["italic"]}]
            collapsed = []
            for r in runs:
                if collapsed and collapsed[-1]["bold"] == r["bold"] and collapsed[-1]["italic"] == r["italic"]:
                    collapsed[-1]["text"] += r["text"]
                else:
                    collapsed.append(dict(r))
            if collapsed and bullet:
                if BULLET_RE.match(collapsed[0]["text"]):
                    collapsed[0]["text"] = BULLET_RE.sub("", collapsed[0]["text"]).strip()
                elif O_BULLET_RE.match(collapsed[0]["text"]):
                    collapsed[0]["text"] = O_BULLET_RE.sub("", collapsed[0]["text"]).strip()
                # ensure a single space after a trailing colon in a bold lead-in
                for i in range(len(collapsed) - 1):
                    if collapsed[i]["text"].rstrip().endswith(":") and collapsed[i+1]["text"] and not collapsed[i+1]["text"].startswith(" "):
                        collapsed[i+1]["text"] = " " + collapsed[i+1]["text"].lstrip()
                if not collapsed[0]["text"]:
                    collapsed[0]["text"] = collapsed[0]["text"].strip()
            run_objs = [_run(r["text"], r["bold"], r["italic"]) for r in collapsed if r["text"].strip()]
            if not run_objs:
                run_objs = [_run(text, merged["bold"], merged["italic"])]
            if level and not bullet:
                items.append((merged["bbox"][1], merged["bbox"][0],
                              Block(kind="heading", level=level, runs=run_objs)))
            elif bullet:
                items.append((merged["bbox"][1], merged["bbox"][0],
                              Block(kind="bullet", level=0, runs=run_objs)))
            else:
                items.append((merged["bbox"][1], merged["bbox"][0],
                              Block(kind="paragraph", runs=run_objs)))

        for y, block in _extract_images(doc, page):
            items.append((y, 0.0, block))

        items.sort(key=lambda t: (round(t[0], 1), t[1]))
        blocks.extend(b for _, _, b in items)

        # crude two-column detector: overlapping lines on both halves of the page
        width = page.rect.width or 1
        left = [l for l in usable if l["bbox"][2] < width * 0.55]
        right = [l for l in usable if l["bbox"][0] > width * 0.45]
        if left and right and len(left) > 3 and len(right) > 3:
            multi_column_pages.append(page_number)

    if scanned_pages:
        warnings.append(
            "Pages with no extractable text: "
            + ", ".join(map(str, scanned_pages[:10]))
            + ("…" if len(scanned_pages) > 10 else "")
            + ". Scanned PDFs need OCR; those pages will be missing from the output.")
    if multi_column_pages:
        warnings.append(
            "Multi-column layout detected on page(s) "
            + ", ".join(map(str, multi_column_pages[:8]))
            + "; text order may need reviewing in the editor.")
    if not blocks:
        warnings.append("No text or tables could be extracted from this PDF.")

    # cover heuristic: first page is a short title-only page
    page_word_counts = [len(''.join(l['text'] for l in pl).split()) for pl in page_lines]
    has_cover = (len(doc) >= 2 and page_word_counts
                 and page_word_counts[0] <= 40
                 and (len(page_word_counts) < 2 or page_word_counts[1] >= page_word_counts[0] * 2))
    model.meta["has_cover"] = has_cover
    heading = next((b for b in blocks if b.kind == "heading" and b.text.strip()), None)
    model.meta["title"] = heading.text.strip() if heading else (
        blocks[0].text.strip()[:120] if blocks and blocks[0].text else name.rsplit(".", 1)[0])
    if heading is not None:
        blocks.remove(heading)
    model.blocks = blocks
    model.warnings = warnings
    return model


def _run(text: str, bold: bool = False, italic: bool = False):
    from .models import Run
    return Run(text=text, bold=bold, italic=italic)
