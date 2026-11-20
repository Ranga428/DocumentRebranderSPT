"""Word document ingestion: body content -> neutral blocks."""
from __future__ import annotations

import io
import re
from datetime import date

import docx
from docx.oxml.ns import qn

from .models import Block, DocumentModel, Run
from .ooxml import ptext, pstyle

BULLET_CHARS = "\u2022\u00b7\u25cf\u25cb\u25a0\u25aa\u25ba\u2023\u2043\uf0b7\uf0a7\uf076\uf0d8"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_KIND_BY_NAME = {
    "title": "title", "subtitle": "subtitle",
    "heading1": "heading1", "heading2": "heading2", "heading3": "heading3",
    "heading4": "heading4", "heading5": "heading5", "heading6": "heading6",
    "quote": "quote", "intensequote": "quote", "caption": "caption",
    "listbullet": "bullet", "listnumber": "number", "listparagraph": "paragraph",
}


def _norm_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def _style_map(doc) -> dict[str, str]:
    out = {}
    for s in doc.styles:
        try:
            out[s.style_id] = s.name
        except Exception:
            continue
    return out


def _num_formats(doc) -> dict[str, str]:
    """numId -> level-0 numFmt (bullet/decimal/...)."""
    try:
        root = doc.part.numbering_part.element
    except Exception:
        return {}
    abstract: dict[str, str] = {}
    for an in root.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        fmt = an.find(qn("w:lvl") + "/" + qn("w:numFmt"))
        abstract[aid] = fmt.get(qn("w:val")) if fmt is not None else "decimal"
    out: dict[str, str] = {}
    for num in root.findall(qn("w:num")):
        aid = num.find(qn("w:abstractNumId"))
        if aid is not None:
            out[num.get(qn("w:numId"))] = abstract.get(aid.get(qn("w:val")), "decimal")
    return out


def _run_text(r) -> str:
    parts = []
    for node in r.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            parts.append("\n")
    return "".join(parts)


def _run(r, link: str | None = None) -> Run | None:
    text = _run_text(r)
    if not text:
        return None
    rpr = r.find(qn("w:rPr"))
    bold = italic = underline = False
    if rpr is not None:
        b = rpr.find(qn("w:b"))
        bold = b is not None and b.get(qn("w:val"), "1") not in ("0", "false")
        i = rpr.find(qn("w:i"))
        italic = i is not None and i.get(qn("w:val"), "1") not in ("0", "false")
        u = rpr.find(qn("w:u"))
        underline = u is not None and u.get(qn("w:val"), "single") != "none"
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None and rs.get(qn("w:val")) in ("Emphasis", "Strong"):
            italic = italic or rs.get(qn("w:val")) == "Emphasis"
            bold = bold or rs.get(qn("w:val")) == "Strong"
    return Run(text=text, bold=bold, italic=italic, underline=underline, link=link)


def _runs(p, part) -> list[Run]:
    out: list[Run] = []

    def walk(node, link=None):
        for child in node:
            tag = child.tag
            if tag == qn("w:r"):
                run = _run(child, link)
                if run:
                    out.append(run)
            elif tag == qn("w:hyperlink"):
                rid = child.get(qn("r:id"))
                url = link
                if rid:
                    rel = part.rels.get(rid)
                    if rel is not None:
                        url = rel.target_ref
                walk(child, url)
            elif tag in (qn("w:ins"), qn("w:smartTag"), qn("w:sdt"),
                         qn("w:sdtContent"), qn("w:fldSimple")):
                walk(child, link)

    walk(p)
    return out


def _style_kind(style_id: str, names: dict[str, str], num_fmt: dict[str, str], p) -> str:
    name = _norm_name(names.get(style_id, style_id or ""))
    kind = _KIND_BY_NAME.get(name, "")
    if not kind:
        if "bullet" in name or "list" in name and "number" not in name:
            kind = "bullet"
        elif "number" in name:
            kind = "number"
    if kind.startswith("heading"):
        return "heading"
    if kind in ("title", "subtitle", "quote", "caption"):
        return kind
    numpr = p.find(qn("w:pPr") + "/" + qn("w:numPr"))
    if numpr is not None:
        nid = numpr.find(qn("w:numId"))
        fmt = num_fmt.get(nid.get(qn("w:val")), "bullet") if nid is not None else "bullet"
        return "number" if "decimal" in fmt or fmt.startswith("upper") or fmt.startswith("lower") else "bullet"
    if kind in ("bullet", "number"):
        return kind
    return "paragraph"


def _heading_level(style_id: str, names: dict[str, str], p) -> int:
    name = _norm_name(names.get(style_id, style_id or ""))
    m = re.search(r"heading(\d)", name)
    if m:
        return max(1, min(6, int(m.group(1))))
    outline = p.find(qn("w:pPr") + "/" + qn("w:outlineLvl"))
    if outline is not None:
        try:
            return max(1, min(6, int(outline.get(qn("w:val"))) + 1))
        except (TypeError, ValueError):
            pass
    return 1


def _cell_text(tc) -> str:
    parts = [ptext(p) for p in tc.findall(qn("w:p"))]
    return "\n".join(x for x in parts if x).strip()


def _table_rows(tbl) -> list[list[str]]:
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        rows.append([_cell_text(tc) for tc in tr.findall(qn("w:tc"))])
    return [r for r in rows if any(c for c in r)]


def _image_blocks(p, doc) -> list[Block]:
    out = []
    for blip in p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        rel = doc.part.rels.get(rid)
        if rel is None or not hasattr(rel.target_part, "blob"):
            continue
        blob = rel.target_part.blob
        extent = None
        for node in p.iter():
            if node.tag.endswith("}extent") and node.get("cx"):
                extent = node
                break
        px = None
        if extent is not None:
            try:
                px = int(int(extent.get("cx")) / 9525)
            except (TypeError, ValueError):
                px = None
        out.append(Block(kind="image", image=blob, width_px=px))
    return out


def _toc_from_sdt(sdt) -> bool:
    xml = "".join(t.text or "" for t in sdt.iter(qn("w:instrText")))
    return "TOC" in xml.upper()


_DOTS = re.compile(r"^(.*?\S)\s*[.·\u2026]{2,}\s*(\d{1,4})\s*$")


def _detect_cover(doc) -> bool:
    first_style = None
    first_text = ""
    seen_first = False
    early_break = False
    count = 0
    for el in doc.element.body:
        if el.tag == qn('w:br') and el.get(qn('w:type')) == 'page':
            if not seen_first or count <= 1:
                early_break = True
            continue
        if el.tag != qn('w:p'):
            continue
        has_pb = any(b.get(qn('w:type')) == 'page' for b in el.iter(qn('w:br')))
        if has_pb:
            if not seen_first or count <= 1:
                early_break = True
        text = ptext(el).strip()
        if not text:
            continue
        if not seen_first:
            first_style = pstyle(el)
            first_text = text
            seen_first = True
            count = 1
            continue
        count += 1
        if count >= 3:
            break
    if not seen_first:
        return False
    style_norm = _norm_name(first_style)
    is_title = (style_norm in ('title', 'heading1', 'heading2')
                or style_norm.startswith('title')
                or (len(first_text.split()) <= 12 and len(first_text) <= 90))
    return bool(is_title and early_break)



def _para_spacing(p) -> dict:
    out = {}
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return out
    sp = ppr.find(qn("w:spacing"))
    if sp is not None:
        def twips(v):
            try:
                return int(v) / 20.0
            except (TypeError, ValueError):
                return None
        if sp.get(qn("w:before")) is not None:
            out["space_before"] = twips(sp.get(qn("w:before")))
        if sp.get(qn("w:after")) is not None:
            out["space_after"] = twips(sp.get(qn("w:after")))
        line = sp.get(qn("w:line"))
        rule = sp.get(qn("w:lineRule"))
        if line is not None:
            try:
                lv = int(line)
            except (TypeError, ValueError):
                lv = None
            if lv is not None:
                if rule == "exact":
                    out["line_spacing"] = lv / 20.0
                elif rule in (None, "auto"):
                    out["line_spacing"] = lv / 240.0
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        def twips(v):
            try:
                return int(v) / 20.0
            except (TypeError, ValueError):
                return None
        if ind.get(qn("w:left")) is not None:
            out["indent_left"] = twips(ind.get(qn("w:left")))
        if ind.get(qn("w:right")) is not None:
            out["indent_right"] = twips(ind.get(qn("w:right")))
        if ind.get(qn("w:firstLine")) is not None:
            out["first_line"] = twips(ind.get(qn("w:firstLine")))
        elif ind.get(qn("w:hanging")) is not None:
            out["first_line"] = -twips(ind.get(qn("w:hanging")))
    return out

def ingest_docx(data: bytes, name: str) -> DocumentModel:
    doc = docx.Document(io.BytesIO(data))
    body = doc.element.body
    names = _style_map(doc)
    num_fmt = _num_formats(doc)
    model = DocumentModel(source_name=name)

    props = doc.core_properties
    title = (props.title or "").strip()
    author = (props.author or "").strip()
    created = props.created or props.modified

    warnings: list[str] = []
    skipped_shapes = 0
    nested_tables = 0

    if title:
        model.meta["title"] = title

    def push(block: Block) -> None:
        if block.kind == "paragraph" and not block.text.strip():
            return
        if block.kind == "title":
            if not model.meta.get("title"):
                model.meta["title"] = block.text.strip()
            else:
                block.kind, block.level = "heading", 1
                model.blocks.append(block)
            return
        if block.kind == "subtitle":
            if not model.meta.get("subtitle"):
                model.meta["subtitle"] = block.text.strip()
                return
            block.kind, block.level = "heading", 2
            model.blocks.append(block)
            return
        model.blocks.append(block)

    def _handle_para(p) -> None:
        nonlocal skipped_shapes
        style_id = pstyle(p)
        kind = _style_kind(style_id, names, num_fmt, p)
        runs = _runs(p, doc.part)

        text = "".join(r.text for r in runs).strip()
        if text:
            bc_count = sum(text.count(bc) for bc in BULLET_CHARS)
            if any(bc in text for bc in BULLET_CHARS) and kind == "paragraph" and bc_count > 1:
                first_bc = next((c for c in BULLET_CHARS if c in text), None)
                lead = text.split(first_bc, 1)[0].strip() if first_bc else ""
                if lead:
                    lb = Block(kind="paragraph", runs=[Run(text=lead)])
                    for k, v in _para_spacing(p).items():
                        setattr(lb, k, v)
                    push(lb)
                rest_text = text.split(first_bc, 1)[1] if first_bc else text
                import re as _re
                segments = _re.split("[" + BULLET_CHARS + "]", rest_text)
                for seg in segments:
                    seg = seg.strip()
                    if seg:
                        bb = Block(kind="bullet", level=0, runs=[Run(text=seg)])
                        for k, v in _para_spacing(p).items():
                            setattr(bb, k, v)
                        push(bb)
            elif any(bc in text for bc in BULLET_CHARS) and kind == "paragraph" and any(text.lstrip().startswith(bc) for bc in BULLET_CHARS):
                bullet = True
                text = text.lstrip(BULLET_CHARS + "-*+").strip()
                block = Block(kind="bullet",
                              level=0,
                              runs=[r for r in runs if r.text])
                if block.runs:
                    block.runs[0].text = block.runs[0].text.lstrip(BULLET_CHARS + "-*+").lstrip()
                if block.text.strip():
                    for k, v in _para_spacing(p).items():
                        setattr(block, k, v)
                    push(block)
            else:
                bullet = (any(text.lstrip().startswith(bc) for bc in BULLET_CHARS)
                          or text.lstrip().startswith("- ") or text.lstrip().startswith("* "))
                if bullet:
                    text = text.lstrip(BULLET_CHARS + "-*+").strip()
                block = Block(kind="heading" if kind.startswith("heading") else ("bullet" if bullet else kind),
                              level=_heading_level(style_id, names, p) if kind.startswith("heading") else 0,
                              runs=[r for r in runs if r.text])
                if bullet and block.runs:
                    block.runs[0].text = block.runs[0].text.lstrip(BULLET_CHARS + "-*+").lstrip()
                if block.kind != "paragraph" and not block.text.strip() and not bullet:
                    block = None
                if block is not None:
                    for k, v in _para_spacing(p).items():
                        setattr(block, k, v)
                    push(block)

        images = _image_blocks(p, doc)
        for img in images:
            model.blocks.append(img)
        if not images and not text and p.find(".//" + qn("w:drawing")) is not None:
            skipped_shapes += 1

    for el in body:
        tag = el.tag
        if tag == qn("w:sectPr"):
            continue

        if tag == qn("w:sdt"):
            is_toc = _toc_from_sdt(el)
            if is_toc:
                continue
            # a content control wrapping ordinary paragraphs
            content = el.find(qn("w:sdtContent"))
            if content is not None:
                for child in content:
                    if child.tag == qn("w:p"):
                        _handle_para(child)
            continue

        if tag == qn("w:p"):
            _handle_para(el)
            continue

        if tag == qn("w:tbl"):
            rows = _table_rows(el)
            nested_tables += len(el.findall(".//" + qn("w:tbl")))
            if rows:
                model.blocks.append(Block(kind="table", rows=rows))
            continue

    # first heading doubles as the title when the document has no title property
    if not model.meta.get("title"):
        first_head = next((b for b in model.blocks if b.kind == "heading"), None)
        if first_head is not None:
            model.meta["title"] = first_head.text.strip()
            model.blocks.remove(first_head)

    model.meta.setdefault("title", title or re.sub(r"[_-]+", " ", name.rsplit(".", 1)[0]).strip())
    if author:
        model.meta["author"] = author
    if created:
        try:
            model.meta["date"] = created.strftime("%d %B %Y")
        except Exception:
            pass

    if skipped_shapes:
        warnings.append(f"{skipped_shapes} drawing-only paragraph(s) had no extractable image "
                        "(likely shaped/text-box content) and were skipped.")
    if nested_tables:
        warnings.append(f"{nested_tables} nested table(s) were flattened or skipped.")
    model.meta["has_cover"] = _detect_cover(doc)
    model.warnings = warnings
    return model


def empty_model(name: str = "") -> DocumentModel:
    return DocumentModel(source_name=name, meta={"date": date.today().strftime("%d %B %Y")})
