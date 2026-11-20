"""Low-level WordprocessingML helpers.

Most of these are lifted from the earlier report converter: they are the part
that is genuinely reusable across document-rebranding projects.
"""
from __future__ import annotations

from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Emu
from docx.text.paragraph import Paragraph

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# --------------------------------------------------------------- inspection
def ptext(el) -> str:
    """Visible text of a paragraph element."""
    if el.tag != qn("w:p"):
        return ""
    return "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()


def pstyle(el) -> str:
    """Style id of a paragraph element ('' for default)."""
    s = el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return s.get(qn("w:val")) if s is not None else ""


def has_drawing(el) -> bool:
    return el.find(".//" + qn("w:drawing")) is not None


def is_blank(el) -> bool:
    return (el.tag == qn("w:p") and not ptext(el) and not has_drawing(el)
            and el.find(".//" + qn("w:sectPr")) is None
            and el.find(".//" + qn("w:bookmarkStart")) is None
            and el.find(".//" + qn("w:br")) is None)


# ------------------------------------------------------------------ editing
def strip_ids(el) -> None:
    """Remove revision/para ids so cloned elements cannot collide."""
    for node in el.iter():
        for attr in list(node.attrib):
            if attr.endswith("}paraId") or attr.endswith("}textId"):
                del node.attrib[attr]


def strip_bookmarks(el) -> None:
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for b in list(el.iter(qn(tag))):
            b.getparent().remove(b)


def set_text(p, text: str) -> None:
    """Replace a paragraph's text, keeping the first text run's formatting."""
    if p.tag != qn("w:p"):
        return
    runs = p.findall(qn("w:r"))
    first = next((r for r in runs if r.find(qn("w:t")) is not None), None)
    if first is None:
        first = runs[0] if runs else OxmlElement("w:r")
        if first.getparent() is None:
            p.append(first)
    for r in runs:
        if r is not first:
            p.remove(r)
    for h in p.findall(qn("w:hyperlink")):
        p.remove(h)
    for child in list(first):
        if child.tag != qn("w:rPr"):
            first.remove(child)
    for i, chunk in enumerate(text.split("\n")):
        if i:
            br = OxmlElement("w:br")
            first.append(br)
        t = OxmlElement("w:t")
        t.set(XML_SPACE, "preserve")
        t.text = chunk
        first.append(t)


def blank_para() -> OxmlElement:
    return OxmlElement("w:p")


def page_break_para() -> OxmlElement:
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def add_page_break(doc) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


# ----------------------------------------------------------------- geometry
def text_width_twips(section) -> int:
    width = Emu(section.page_width - section.left_margin - section.right_margin)
    return int(width.twips)


def usable_emu(section) -> tuple[int, int]:
    w = Emu(section.page_width - section.left_margin - section.right_margin)
    h = Emu(section.page_height - section.top_margin - section.bottom_margin)
    return int(w.emu), int(h.emu)


# ------------------------------------------------------------------ numbering
def ensure_numbering(doc) -> dict[str, int]:
    """Create (once) bullet and decimal list definitions. Returns num ids."""
    numbering = doc.part.numbering_part.element
    existing = {n.get(qn("w:abstractNumId")): n
                for n in numbering.findall(qn("w:abstractNum"))}
    nums = {n.get(qn("w:numId")) for n in numbering.findall(qn("w:num"))}

    def make_abstract(abstract_id: int, fmt: str, texts: list[str]) -> None:
        levels = []
        for ilvl, text in enumerate(texts):
            left = 720 + 360 * ilvl
            levels.append(
                f'<w:lvl w:ilvl="{ilvl}">'
                '<w:start w:val="1"/>'
                f'<w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{text}"/>'
                '<w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Onest" w:hAnsi="Onest" w:cs="Onest"/></w:rPr>'
                '</w:lvl>'
            )
        abstract = parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
            '<w:multiLevelType w:val="hybridMultilevel"/>'
            + "".join(levels) + '</w:abstractNum>')
        first_num = numbering.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            numbering.append(abstract)

    def make_num(num_id: int, abstract_id: int) -> None:
        num = parse_xml(
            f'<w:num {nsdecls("w")} w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abstract_id}"/></w:num>')
        numbering.append(num)

    if "900" not in existing:
        make_abstract(900, "bullet", ["\u2022", "\u25e6", "\u25aa", "\u2022", "\u25e6"])
    if "900" not in nums:
        make_num(900, 900)
    if "901" not in existing:
        make_abstract(901, "decimal", [f"%{i}." for i in range(1, 6)])
    if "901" not in nums:
        make_num(901, 901)
    return {"bullet": 900, "number": 901}


def new_num(doc, proto_num_id: int) -> int:
    """Clone a numbering definition so a list can restart at 1."""
    numbering = doc.part.numbering_part.element
    nums = numbering.findall(qn("w:num"))
    proto = next((n for n in nums
                  if n.get(qn("w:numId")) == str(proto_num_id)), None)
    if proto is None:
        return proto_num_id
    abs_id = proto.find(qn("w:abstractNumId")).get(qn("w:val"))
    used = [int(n.get(qn("w:numId"))) for n in nums
            if (n.get(qn("w:numId")) or "").isdigit()]
    new_id = max(used + [proto_num_id]) + 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_id))
    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), abs_id)
    num.append(abstract)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:startOverride")
    start.set(qn("w:val"), "1")
    override.append(start)
    num.append(override)
    numbering.append(num)
    return new_id


def apply_list(paragraph: Paragraph, num_id: int, level: int = 0) -> None:
    p = paragraph._p
    ppr = p.get_or_add_pPr()
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        numpr = OxmlElement("w:numPr")
        ppr.append(numpr)
    ilvl = numpr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        numpr.append(ilvl)
    ilvl.set(qn("w:val"), str(level))
    numid = numpr.find(qn("w:numId"))
    if numid is None:
        numid = OxmlElement("w:numId")
        numpr.append(numid)
    numid.set(qn("w:val"), str(num_id))
    # numbering already carries the indent; a first-line indent breaks it
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        ppr.remove(ind)


# --------------------------------------------------------------------- fonts
def force_font(doc, font: str = "Onest") -> None:
    """Make every run in the package use *font*, theme fonts included."""
    roots = [doc.element, doc.styles.element]
    try:
        roots.append(doc.part.numbering_part.element)
    except Exception:
        pass
    for sec in doc.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            try:
                roots.append(part._element)
            except Exception:
                pass

    for root in roots:
        for rf in root.iter(qn("w:rFonts")):
            for attr in list(rf.attrib):
                del rf.attrib[attr]
            for key in ("ascii", "hAnsi", "cs", "eastAsia"):
                rf.set(qn("w:" + key), font)

    # document defaults
    rpr_default = doc.styles.element.find(
        qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
    if rpr_default is not None:
        rf = rpr_default.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr_default.append(rf)
        for key in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(qn("w:" + key), font)
