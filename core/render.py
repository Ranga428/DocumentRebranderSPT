"""Render a DocumentModel into the fixed brand template."""
from __future__ import annotations

import copy
import io
import re

import docx
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER)
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor, Twips

from .models import Block, DocumentModel, Run
from .ooxml import (apply_list, ensure_numbering, force_font, new_num, set_text,
                    strip_bookmarks, strip_ids, text_width_twips, usable_emu)
from .template import manifest

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# --------------------------------------------------------------------- runs
def _run_element(run: Run, hyperlink: bool = False) -> OxmlElement:
    r = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    if run.bold:
        props.append(OxmlElement("w:b"))
    if run.italic:
        props.append(OxmlElement("w:i"))
    if run.underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        props.append(u)
    if hyperlink:
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "0563C1")
        props.append(color)
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        props.append(u)
    if len(props):
        r.append(props)
    for i, chunk in enumerate((run.text or "").split("\n")):
        if i:
            r.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.set(XML_SPACE, "preserve")
        t.text = chunk
        r.append(t)
    return r


def _write_runs(paragraph, runs: list[Run]) -> None:
    for run in runs:
        if run.text == "":
            continue
        if run.link:
            rid = paragraph.part.relate_to(run.link, RT.HYPERLINK, is_external=True)
            link = OxmlElement("w:hyperlink")
            link.set(qn("r:id"), rid)
            link.append(_run_element(run, hyperlink=True))
            paragraph._p.append(link)
        else:
            paragraph._p.append(_run_element(run))


# ------------------------------------------------------------------ helpers
class _Renderer:
    def __init__(self, doc, model: DocumentModel, *, cover: bool, control: str,
                 table_proto=None):
        self.d = doc
        self.model = model
        self.cover = cover
        self.control = control
        self.warnings: list[str] = []
        self.section = doc.sections[0]
        self.body = doc.element.body
        self.sect_pr = self.body.find(qn("w:sectPr"))
        self.styles = {s.style_id: s for s in doc.styles}
        self.nums = ensure_numbering(doc)
        self.width_twips = text_width_twips(self.section)
        self.table_proto = table_proto
        self.heading_order = self._heading_order()
        self.number_num = self.nums["number"]
        self.prev_kind = ""
        levels = [b.level or 1 for b in model.blocks if b.kind == "heading"]
        self.min_heading = min(levels) if levels else 1

    # -- infrastructure ---------------------------------------------------
    def _style_size(self, style) -> float:
        sz = style.element.find(qn("w:rPr") + "/" + qn("w:sz"))
        if sz is not None:
            try:
                return int(sz.get(qn("w:val"))) / 2
            except (TypeError, ValueError):
                pass
        if style.font.size is not None:
            return float(style.font.size.pt)
        return 11.0

    def _heading_order(self) -> list[str]:
        """Heading style ids ordered by visual prominence.

        The brand template's Heading 1 is 11pt bold while Heading 2 is 16pt,
        so mapping document depth to style *numbers* would invert the
        hierarchy. Map level 1 to the largest heading style instead.
        """
        ranked = []
        for style_id, style in self.styles.items():
            if not re.fullmatch(r"Heading[1-6]", style_id):
                continue
            bold = style.element.find(qn("w:rPr") + "/" + qn("w:b")) is not None
            ranked.append((-self._style_size(style), 0 if bold else 1, style_id))
        ranked.sort()
        return [style_id for _, _, style_id in ranked]

    def style(self, style_id: str):
        return self.styles.get(style_id)

    def para(self, style_id: str = "Normal", runs: list[Run] | None = None,
             text: str | None = None):
        p = self.d.add_paragraph(style=self.style(style_id))
        if runs:
            _write_runs(p, runs)
        elif text is not None:
            p.add_run(text)
        return p

    def _apply_spacing(self, p, block: Block) -> None:
        from docx.shared import Pt
        if block.space_before is not None:
            p.paragraph_format.space_before = Pt(block.space_before)
        if block.space_after is not None:
            p.paragraph_format.space_after = Pt(block.space_after)
        if block.line_spacing is not None:
            p.paragraph_format.line_spacing = block.line_spacing
        if block.indent_left is not None:
            p.paragraph_format.left_indent = Pt(block.indent_left)
        if block.indent_right is not None:
            p.paragraph_format.right_indent = Pt(block.indent_right)
        if block.first_line is not None:
            p.paragraph_format.first_line_indent = Pt(block.first_line)

    def page_break(self) -> None:
        p = self.d.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)

    # -- cover / toc ------------------------------------------------------
    def build_cover(self) -> None:
        meta = self.model.meta
        title = str(meta.get("title") or "").strip()
        subtitle = str(meta.get("subtitle") or "").strip()
        meta_bits = [str(meta.get("date") or "").strip()]
        if self.control:
            meta_bits.append(f"Control No. {self.control}")
        if meta.get("author"):
            meta_bits.append(str(meta["author"]).strip())
        meta_bits = [m for m in meta_bits if m]

        # The first-page header carries a 21 x 5 cm anchored logo whose bottom
        # sits about 3 cm into the text area, so the cover has to start below it.
        for _ in range(6):
            self.d.add_paragraph()
        if title:
            p = self.para("Title", text=title)
            p.paragraph_format.space_before = Pt(18)
        if subtitle:
            self.para("Subtitle", text=subtitle)
        if meta_bits:
            p = self.para("Normal", text="   \u00b7   ".join(meta_bits))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        self.page_break()

    # -- content ----------------------------------------------------------
    def add_heading(self, block: Block) -> None:
        level = max(1, min(6, block.level or 1))
        # normalise: the shallowest heading still in the body is level 1
        effective = max(1, level - self.min_heading + 1)
        if self.heading_order:
            style_id = self.heading_order[min(effective - 1, len(self.heading_order) - 1)]
        else:
            style_id = f"Heading{min(effective, 6)}"
        p = self.para(style_id, runs=block.runs)
        self._apply_spacing(p, block)

    def add_paragraph(self, block: Block) -> None:
        if block.kind in ("bullet", "number"):
            # ensure bold on "Label:" lead-in (up to first colon) + space after colon
            runs = []
            for r in block.runs:
                m = re.match(r"^([^:\\n]{1,45}):(.*)$", r.text)
                if m and not r.bold:
                    prefix = m.group(1) + ":"
                    rest = m.group(2)
                    runs.append(Run(text=prefix, bold=True, italic=r.italic,
                                    underline=r.underline, link=r.link))
                    if rest.strip():
                        sep = "" if rest.startswith(" ") else " "
                        runs.append(Run(text=sep + rest, bold=r.bold,
                                        italic=r.italic, underline=r.underline, link=r.link))
                else:
                    runs.append(r)
            # collapse soft line breaks into spaces so text wraps naturally
            runs = [Run(text=r.text.replace("\n", " "), bold=r.bold, italic=r.italic,
                        underline=r.underline, link=r.link) for r in runs]
            p = self.para("Normal", runs=runs)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if block.kind == "number":
                if self.prev_kind != "number":
                    self.number_num = new_num(self.d, self.nums["number"])
                num_id = self.number_num
            else:
                num_id = self.nums["bullet"]
            apply_list(p, num_id, min(block.level, 4))
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            self._apply_spacing(p, block)
            return
        if block.kind == "quote":
            p = self.para("Normal", runs=block.runs)
            p.paragraph_format.left_indent = Twips(720)
            for run in p.runs:
                run.italic = True
            self._apply_spacing(p, block)
            return
        if block.kind == "caption":
            p = self.para("Normal", runs=block.runs)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.italic = True
                run.font.size = Pt(9)
            self._apply_spacing(p, block)
            return
        p = self.para("Normal", runs=block.runs)
        self._apply_spacing(p, block)

    def add_hr(self) -> None:
        p = self.d.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        pbdr.append(bottom)
        pPr.append(pbdr)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

    def add_image(self, block: Block) -> None:
        if not block.image:
            return
        max_w, max_h = usable_emu(self.section)
        width = min(block.width_px * 9525 if block.width_px else max_w, max_w)
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            shape = p.add_run().add_picture(io.BytesIO(block.image), width=Emu(int(width)))
        except Exception as exc:
            self.warnings.append(f"One image could not be inserted: {exc}")
            self.body.remove(p._p)
            return
        if shape.height > max_h * 0.75:
            factor = (max_h * 0.75) / shape.height
            shape.height = int(shape.height * factor)
            shape.width = int(shape.width * factor)
        self.d.add_paragraph()

    def add_table(self, block: Block) -> None:
        if not block.rows:
            return
        if self.table_proto is None:
            self.warnings.append("Template has no table prototype; tables were skipped.")
            return
        ncols = max(len(r) for r in block.rows)
        rows = [list(r) + [""] * (ncols - len(r)) for r in block.rows]
        colw = self.width_twips // max(1, ncols)

        tbl = copy.deepcopy(self.table_proto)
        strip_ids(tbl)
        strip_bookmarks(tbl)
        trs = tbl.findall(qn("w:tr"))
        if not trs:
            return
        header_proto = copy.deepcopy(trs[0])
        body_proto = copy.deepcopy(trs[1] if len(trs) > 1 else trs[0])
        for tr in trs:
            tbl.remove(tr)

        tbl_pr = tbl.find(qn("w:tblPr"))
        grid = tbl.find(qn("w:tblGrid"))
        if grid is not None:
            tbl.remove(grid)
        grid = OxmlElement("w:tblGrid")
        for _ in range(ncols):
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(colw))
            grid.append(gc)
        tbl_pr.addnext(grid)

        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:w"), str(self.width_twips))
        tbl_w.set(qn("w:type"), "dxa")

        for ri, row in enumerate(rows):
            proto_tr = header_proto if (block.header_row and ri == 0) else body_proto
            tr = copy.deepcopy(proto_tr)
            strip_ids(tr)
            strip_bookmarks(tr)
            cells = tr.findall(qn("w:tc"))
            cell_proto = copy.deepcopy(cells[0]) if cells else None
            for cell in cells:
                tr.remove(cell)
            if cell_proto is None:
                continue
            for ci in range(ncols):
                tc = copy.deepcopy(cell_proto)
                strip_ids(tc)
                strip_bookmarks(tc)
                tc_w = tc.find(qn("w:tcPr") + "/" + qn("w:tcW"))
                if tc_w is not None:
                    tc_w.set(qn("w:w"), str(colw))
                    tc_w.set(qn("w:type"), "dxa")
                paragraphs = tc.findall(qn("w:p"))
                for extra in paragraphs[1:]:
                    tc.remove(extra)
                p = paragraphs[0] if paragraphs else OxmlElement("w:p")
                if not paragraphs:
                    tc.append(p)
                set_text(p, row[ci] if ci < len(row) else "")
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                tr.append(tc)
            if block.header_row and ri == 0:
                tr_pr = tr.find(qn("w:trPr"))
                if tr_pr is None:
                    tr_pr = OxmlElement("w:trPr")
                    tr.insert(0, tr_pr)
                repeat = OxmlElement("w:tblHeader")
                repeat.set(qn("w:val"), "1")
                tr_pr.append(repeat)
            tbl.append(tr)

        self.sect_pr.addprevious(tbl)
        self.d.add_paragraph()

    # -- document ---------------------------------------------------------
    def build(self) -> list[str]:
        # safety net: split any paragraph that still contains inline bullet chars
        try:
            from core.ingest import _split_inline_bullets, _split_numbered_headings
            self.model.blocks = _split_inline_bullets(self.model.blocks)
            self.model.blocks = _split_numbered_headings(self.model.blocks)
        except Exception:
            pass
        # second safety net: any paragraph whose text starts with a non-alphanumeric
        # symbol is treated as a bullet (handles unknown bullet glyphs)
        try:
            import unicodedata
            from core.models import Block, Run
            new_blocks = []
            for b in self.model.blocks:
                if b.kind == "paragraph" and b.text.strip():
                    first = b.text.strip()[0]
                    cat = unicodedata.category(first)
                    if (cat.startswith("P") or cat.startswith("S")) and first not in (".", ",", "!", "?", ":", ";"):
                        txt = b.text.strip()[1:].lstrip()
                        if txt:
                            nb = Block(kind="bullet", level=b.level, runs=[Run(text=txt)])
                            new_blocks.append(nb)
                            continue
                new_blocks.append(b)
            self.model.blocks = new_blocks
        except Exception:
            pass
        self.apply_control_number()
        if not self.cover:
            # The big first-page logo belongs to the cover. Without a cover,
            # let page 1 use the normal (small) header so text cannot overlap
            # the anchored image.
            self.section.different_first_page_header_footer = False
        if self.cover:
            self.build_cover()
        elif self.model.title:
            self.para("Title", text=self.model.title)
            if self.model.subtitle:
                self.para("Subtitle", text=self.model.subtitle)

        for block in self.model.blocks:
            if block.kind == "title":
                if not self.cover and self.model.title:
                    continue
                p = self.para("Title", runs=block.runs)
                self._apply_spacing(p, block)
            elif block.kind == "subtitle":
                if not self.cover and self.model.subtitle:
                    continue
                p = self.para("Subtitle", runs=block.runs)
                self._apply_spacing(p, block)
            elif block.kind == "heading":
                self.add_heading(block)
            elif block.kind == "table":
                self.add_table(block)
            elif block.kind == "image":
                self.add_image(block)
            elif block.kind == "hr":
                self.add_hr()
            elif block.kind == "pagebreak":
                self.page_break()
            else:
                self.add_paragraph(block)
            self.prev_kind = block.kind
        return self.warnings

    def apply_control_number(self) -> None:
        for sec in self.d.sections:
            for part in (sec.header, sec.first_page_header, sec.footer,
                         sec.first_page_footer):
                try:
                    element = part._element
                except Exception:
                    continue
                for p in element.iter(qn("w:p")):
                    text = "".join(t.text or "" for t in p.iter(qn("w:t")))
                    if "CONTROL#" in text:
                        set_text(p, f"CONTROL#{self.control}" if self.control else "")


def render(model: DocumentModel, template_bytes: bytes, *, cover: bool = True,
           control_number: str = "") -> tuple[bytes, list[str]]:
    """Return (docx bytes, warnings)."""
    doc = docx.Document(io.BytesIO(template_bytes))
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    # the demo table is the brand table prototype; capture it before clearing
    first_tbl = body.find(qn("w:tbl"))
    table_proto = copy.deepcopy(first_tbl) if first_tbl is not None else None

    # discard the demo body, keep section properties (headers/footers live there)
    for el in list(body):
        if el is not sect_pr:
            body.remove(el)

    renderer = _Renderer(doc, model, cover=cover, control=control_number or "",
                         table_proto=table_proto)
    warnings = renderer.build()

    force_font(doc, manifest().get("font", "Onest"))
    doc.core_properties.title = model.title or "Document"
    if model.meta.get("author"):
        doc.core_properties.author = str(model.meta["author"])
    doc.core_properties.comments = "Rebranded with Document Rebrander"

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), warnings


def apply_control_number_bytes(docx_bytes, control_number):
    import io as _io
    doc = docx.Document(_io.BytesIO(docx_bytes))
    from .ooxml import set_text
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header, sec.footer,
                     sec.first_page_footer):
            try:
                element = part._element
            except Exception:
                continue
            for pp in element.iter(qn("w:p")):
                text = "".join(t.text or "" for t in pp.iter(qn("w:t")))
                if "CONTROL#" in text:
                    set_text(pp, "CONTROL#" + control_number if control_number else "")
    out = _io.BytesIO()
    doc.save(out)
    return out.getvalue()
