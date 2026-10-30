"""VAPT report converter: original PDF  ->  branded Word report.

The converted sample .docx is used as the template. Fixed wording stays in the
template; variable data (cover, dates, counts, nmap, scope, findings, TOC) is
parsed from the original PDF. Rewritten finding text is learned from the
template itself (title -> details/steps) and re-applied to matching titles.
"""
import copy
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import warnings

warnings.filterwarnings("ignore")
import docx
import fitz  # PyMuPDF
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

LABELS = ["Description:", "Corrective Actions:", "Reference:", "Details of finding:",
          "Affected Hosts:", "Steps to Reproduce:"]
FIELD = {"Description:": "description", "Corrective Actions:": "corrective",
         "Reference:": "reference", "Details of finding:": "details",
         "Affected Hosts:": "hosts", "Steps to Reproduce:": "steps"}
DEFAULT_CFG = {
    "documented_by": "Jenel Esteron",
    "asset_label": None,            # None -> first word of original cover title, lower-cased
    "text_patches": [["*.foo.example.org", "*.sample.example.org"]],
}
NUMWORDS = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"]


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


# ----------------------------------------------------------------- PDF parsing
def _lines(doc):
    """Return per-page line lists [(text, y0, y1)] with running header/footer removed."""
    company = doc[1].get_text().strip().splitlines()[0].strip() if len(doc) > 1 else ""
    pages = []
    for p in doc:
        out = []
        for b in p.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for ln in b["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).rstrip()
                if not t.strip():
                    continue
                if t.strip() == company or re.fullmatch(r"Page \d+", t.strip()) \
                        or re.fullmatch(r"www\.\S+", t.strip()) or t.strip() == "...":
                    continue
                out.append((t.strip(), ln["bbox"][1], ln["bbox"][3]))
        out.sort(key=lambda x: x[1])
        pages.append(out)
    return pages, company


def _join(lines):
    """Join wrapped lines; glue URL fragments and hyphen-wrapped words."""
    s = ""
    for i, t in enumerate(lines):
        if not s:
            s = t
            continue
        last = s.split(" ")[-1]
        if last.startswith("http") and " " not in t.split("://")[0] and not t[:1].isupper():
            s += t
        elif s.endswith("-") and t[:1].islower():
            s += t
        else:
            s += " " + t
    return s.strip()


def _paras(items):
    """items: [(text,y0,y1,page)] -> paragraph strings using vertical gaps."""
    groups, cur, prev = [], [], None
    for t, y0, y1, pg in items:
        if prev and prev[3] == pg and (y0 - prev[2]) > 0.5 * (y1 - y0):
            groups.append(cur)
            cur = []
        cur.append(t)
        prev = (t, y0, y1, pg)
    if cur:
        groups.append(cur)
    return [_join(g) for g in groups]


def _steps(items):
    out = []
    for t, *_ in items:
        m = re.match(r"^\d+\.\s+(.*)", t)
        if m or not out:
            out.append([m.group(1) if m else t])
        else:
            out[-1].append(t)
    return [_join(x) for x in out]


def parse_original(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages, company = _lines(doc)
    full = ["\n".join(t for t, *_ in pg) for pg in pages]
    rep = {"company_old": company, "warnings": []}

    # cover
    cover = [t for t, *_ in pages[0]]
    rep["cover_title"] = cover[0] if cover else ""
    m = re.search(r"ASSESSMENT ID:\s*(\S+)", full[0])
    rep["assessment_id"] = m.group(1) if m else ""

    alltxt = "\n".join(full)
    m = re.search(r"Created:\s*(.+)", alltxt)
    rep["created"] = m.group(1).strip() if m else ""
    m = re.search(r"Reference:\s*(\S+)", alltxt)
    rep["reference"] = m.group(1) if m else rep["assessment_id"]

    # executive summary dates
    m = re.search(r"from (\w+ \d+, \d{4}),? to (\w+ \d+, \d{4})", alltxt)
    rep["start"], rep["end"] = (m.group(1), m.group(2)) if m else ("", "")

    # overview & scope
    first = [f.split("\n")[0].strip() if f else "" for f in full]
    ov = next((i for i, f in enumerate(first) if f == "Overview And Scope"), None)
    rep["dates"], rep["scope"] = [], []
    if ov is not None:
        ls = [t for t, *_ in pages[ov]]
        sec = None
        for t in ls:
            if t.startswith("Test Dates"):
                sec = "d"
            elif t.startswith("Test Scope"):
                sec = "s"
            elif t == "Overview And Scope":
                continue
            elif sec == "d":
                rep["dates"].append(t)
            elif sec == "s":
                rep["scope"].append(t)
    # recon block
    ex = next((i for i, f in enumerate(first) if f == "Executive Summary"), None)
    rep["recon"] = []
    if ex is not None and ov is not None:
        on = False
        for i in range(ex, ov):
            for t, *_ in pages[i]:
                if re.match(r"Server-level Reconnaissance", t, re.I):
                    on = True
                    continue
                if on:
                    rep["recon"].append(t)

    # findings
    pt = next((i for i, f in enumerate(first) if f == "Penetration Test Results"), None)
    ap = next((i for i, f in enumerate(first) if f.startswith("Appendix A")), len(doc))
    findings, cur, field = [], None, None
    xref_pages = {}
    for pi, p in enumerate(doc):
        for im in p.get_image_info(xrefs=True):
            xref_pages.setdefault(im["xref"], set()).add(pi)
    if pt is not None:
        for pi in range(pt, ap):
            for im in sorted(doc[pi].get_image_info(xrefs=True), key=lambda x: x["bbox"][1]):
                b = im["bbox"]
                w, h = b[2] - b[0], b[3] - b[1]
                if w > 100 and h > 40 and not (w > 590 and h > 830) and len(xref_pages[im["xref"]]) < 3 and cur is not None:
                    pix = fitz.Pixmap(doc, im["xref"])
                    if pix.n - pix.alpha > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    cur["images"].append(pix.tobytes("png"))
            for t, y0, y1 in pages[pi]:
                m = re.match(r"^Findings:\s*(\S+)\s*-\s*(\d+)", t)
                if m:
                    cur = {"id": f"{m.group(1)} - {m.group(2)}", "cvss": "", "severity": "", "status": "",
                           "title": [], "images": [], "raw": {k: [] for k in FIELD.values()}}
                    findings.append(cur)
                    field = "head"
                    continue
                if cur is None:
                    continue
                if field == "head":
                    if t.startswith("CVSS:"):
                        cur["cvss"] = t[5:].strip()
                    elif t.startswith("SEVERITY:"):
                        cur["severity"] = t[9:].strip()
                    elif t.startswith("STATUS:"):
                        cur["status"] = t[7:].strip()
                        field = "title"
                    continue
                if t in LABELS:
                    field = FIELD[t]
                    continue
                if field == "title":
                    cur["title"].append(t)
                elif field:
                    cur["raw"][field].append((t, y0, y1, pi))
    for f in findings:
        f["title"] = " ".join(f["title"]).strip()
        r = f.pop("raw")
        f["description"] = _paras(r["description"])
        f["corrective"] = _paras(r["corrective"])
        f["reference"] = _paras(r["reference"])
        f["details"] = _paras(r["details"])
        f["hosts"] = [x for x in (t for t, *_ in r["hosts"])]
        f["steps"] = _steps(r["steps"])
    rep["findings"] = findings
    counts = {k: 0 for k in ("Critical", "High", "Medium", "Low")}
    for f in findings:
        k = f["severity"].strip().capitalize()
        if k in counts:
            counts[k] += 1
    rep["counts"] = counts
    m = re.findall(r"^(Critical|High|Medium|Low):\s*(\d+)", alltxt, re.M)
    if m and {k: int(v) for k, v in m} != counts:
        rep["warnings"].append(f"Severity counts in the original summary {dict(m)} differ from the findings found ({counts}).")
    return rep


# ------------------------------------------------------------ rewrite library
def learn_library(template_bytes):
    """Learn title -> reworded details/steps from a converted .docx."""
    d = docx.Document(io.BytesIO(template_bytes))
    lib, cur, field = {}, None, None
    for el in d.element.body:
        tag = el.tag.split("}")[1]
        if tag == "tbl" and "Vulnerability ID" in "".join(el.itertext()):
            hd = [Paragraph(p, d) for p in el.iter(qn("w:p"))]
            title = next((p.text.strip() for p in hd if p.style.name == "Heading 2"), "")
            cur = lib.setdefault(norm(title), {"title": title, "details": [], "steps": []})
            field = None
            continue
        if tag != "p" or cur is None:
            continue
        p = Paragraph(el, d)
        if p.style.name.startswith("Heading 1") and p.text.strip():
            cur = None
            continue
        t = p.text.strip()
        if p.style.name == "Heading 2" and t:
            field = FIELD.get(t if t.endswith(":") else t + ":") or FIELD.get(t)
            continue
        if field in ("details", "steps") and t:
            cur[field].append(t)
    return lib


# ------------------------------------------------------------------- the plan
def make_plan(rep, lib, cfg=None):
    cfg = {**DEFAULT_CFG, **(cfg or {})}
    plan = {"cfg": cfg, "warnings": list(rep["warnings"]), "review": []}
    plan["asset_label"] = cfg["asset_label"] or (rep["cover_title"].split() or ["web3"])[0].lower()
    plan["url"] = rep["scope"][0] if rep["scope"] else ""
    for k in ("assessment_id", "created", "reference", "start", "end", "dates", "scope", "recon", "counts"):
        plan[k] = rep[k]
    plan["findings"] = []
    for f in rep["findings"]:
        g = {k: f[k] for k in ("id", "cvss", "severity", "status", "title", "images")}
        patched = False
        for k in ("description", "corrective", "reference"):
            out = []
            for para in f[k]:
                n = para
                for a, b in cfg["text_patches"]:
                    n = n.replace(a, b)
                patched |= n != para
                out.append(n)
            g[k] = out
        g["hosts"] = [h if re.match(r"https?://", h) else f"https://{h.strip('/')}/" for h in f["hosts"]]
        e = lib.get(norm(f["title"]))
        for k in ("details", "steps"):
            if e and e[k]:
                g[k] = list(e[k])
                src = "template wording" + ("" if e[k] == f[k] else " (replaced original text)")
            else:
                g[k] = list(f[k])
                src = "copied from original - NO reworded version known, please review"
                plan["warnings"].append(f"{f['title']}: no reworded '{k}' in the template library; original wording used.")
            plan["review"].append((f["title"], k, src))
        plan["review"].append((f["title"], "description/corrective/reference",
                               "copied" + (" with text patches" if patched else "")))
        # consistency check: header value quoted in details vs steps
        dv = re.search(r"X-Frame-Options:?\s*(SAMEORIGIN|DENY)", " ".join(g["details"]))
        for i, s in enumerate(g["steps"]):
            sv = re.search(r"set to (SAMEORIGIN|DENY)", s)
            if dv and sv and sv.group(1) != dv.group(1):
                g["steps"][i] = s.replace(sv.group(0), f"set to {dv.group(1)}")
                plan["warnings"].append(f"{f['title']}: step {i + 1} said {sv.group(1)} but the details say {dv.group(1)}; step corrected.")
        plan["findings"].append(g)
    if not plan["findings"]:
        plan["warnings"].append("No findings were detected in the PDF.")
    plan["images_missing"] = [g["title"] for g in plan["findings"] if not g["images"]]
    return plan


# ------------------------------------------------------------------- docx edit
def _runs(p):
    return [r for r in p.findall(qn("w:r"))]


def set_text(p, text):
    """Replace paragraph text, keeping the first text run's formatting."""
    runs = _runs(p)
    first = next((r for r in runs if r.find(qn("w:t")) is not None), None)
    if first is None:
        first = runs[0] if runs else None
    if first is None:
        first = OxmlElement("w:r")
        p.append(first)
    for r in runs:
        if r is not first:
            p.remove(r)
    for h in p.findall(qn("w:hyperlink")):
        p.remove(h)
    for c in list(first):
        if c.tag != qn("w:rPr"):
            first.remove(c)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    first.append(t)


def ptext(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t"))).strip() if el.tag == qn("w:p") else ""


def pstyle(el):
    s = el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return s.get(qn("w:val")) if s is not None else ""


def strip_bookmarks(el):
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for b in list(el.iter(qn(tag))):
            b.getparent().remove(b)


def page_break_before(el):
    Paragraph(el, None).paragraph_format.page_break_before = True


def spacer():
    """Empty paragraph that starts a new page (gives the header some air)."""
    p = OxmlElement("w:p")
    p.append(OxmlElement("w:pPr"))
    Paragraph(p, None).paragraph_format.page_break_before = True
    return p


def blank():
    return OxmlElement("w:p")


def is_blank(el):
    return (el.tag == qn("w:p") and not ptext(el) and el.find(".//" + qn("w:drawing")) is None
            and el.find(".//" + qn("w:sectPr")) is None and el.find(".//" + qn("w:bookmarkStart")) is None
            and el.find(".//" + qn("w:br")) is None)


def new_num(d, proto_num_id):
    numbering = d.part.numbering_part.element
    nums = numbering.findall(qn("w:num"))
    abs_id = next(n for n in nums if n.get(qn("w:numId")) == str(proto_num_id)).find(qn("w:abstractNumId")).get(qn("w:val"))
    new_id = max(int(n.get(qn("w:numId"))) for n in nums) + 1
    n = OxmlElement("w:num")
    n.set(qn("w:numId"), str(new_id))
    a = OxmlElement("w:abstractNumId")
    a.set(qn("w:val"), abs_id)
    n.append(a)
    ov = OxmlElement("w:lvlOverride")
    ov.set(qn("w:ilvl"), "0")
    so = OxmlElement("w:startOverride")
    so.set(qn("w:val"), "1")
    ov.append(so)
    n.append(ov)
    numbering.append(n)
    return new_id


def build_docx(plan, template_bytes, update_toc=True):
    """Return (docx_bytes, warnings)."""
    warns = list(plan["warnings"])
    d = docx.Document(io.BytesIO(template_bytes))
    body = d.element.body
    els = list(body)

    def find(pred, start=0):
        for i in range(start, len(els)):
            if pred(els[i]):
                return i
        return None

    def find_text(prefix, style=None, start=0):
        return find(lambda e: e.tag == qn("w:p") and ptext(e).startswith(prefix)
                    and (style is None or pstyle(e) == style), start)

    # ---- cover / copyright
    i = find(lambda e: e.tag == qn("w:p") and pstyle(e) == "Heading1" and "http" in ptext(e))
    p = els[i]
    rs = [r for r in _runs(p) if r.find(qn("w:t")) is not None]
    rs[0].find(qn("w:t")).text = f"{plan['asset_label']} | "
    rs[1].find(qn("w:t")).text = plan["url"] + " "
    for r in rs[2:]:
        p.remove(r)
    set_text(els[find_text("ASSESSMENT ID")], f"ASSESSMENT ID: {plan['assessment_id']}")
    set_text(els[find_text("Created:")], f"Created: {plan['created']}")
    set_text(els[find_text("Reference:")], f"Reference: {plan['reference']}")
    set_text(els[find_text("Documented By:")], f"Documented By: {plan['cfg']['documented_by']}")

    # cover: move the lower "CONFIDENTIAL DOCUMENT" up by two lines
    conf = [e for e in els if e.tag == qn("w:p") and ptext(e) == "CONFIDENTIAL DOCUMENT"]
    if len(conf) >= 2:
        prev, removed = conf[1].getprevious(), 0
        while prev is not None and removed < 2 and is_blank(prev):
            nxt = prev.getprevious()
            body.remove(prev)
            prev, removed = nxt, removed + 1

    # ---- executive summary
    i = find_text("Executive Summary", "Heading1")
    j = find_text("Sapient Technology, Inc. conducted", None, i)
    p = els[j]
    ts = [r for r in _runs(p) if r.find(qn("w:t")) is not None]
    ts[1].find(qn("w:t")).text = plan["url"]
    t2 = ts[2].find(qn("w:t"))
    t2.text = re.sub(r"from .+? to .+?\. Detailed", f"from {plan['start']}, to {plan['end']}. Detailed", t2.text, 1)
    c = plan["counts"]
    for k in ("Critical", "High", "Medium", "Low"):
        set_text(els[find_text(f"{k}:", None, i)], f"{k}: {c[k]} ")
    total = sum(c.values())
    k = find_text("The assessment was successfully conducted", None, i)
    if c["Critical"] or c["High"]:
        warns.append("Executive Summary wording says the target showed 'a fair level of security'; "
                     "check it still fits with Critical/High findings.")
    k = find_text("The assessment did not result", None, i)
    if k is not None:
        pe = els[k]
        t = ptext(pe)
        if total and c["Low"] == total:
            n = NUMWORDS[total] if total < len(NUMWORDS) else str(total)
            new = re.sub(r"three low-severity findings", f"{n} low-severity finding" + ("" if total == 1 else "s"), t)
            if total == 1:
                new = new.replace("were identified", "was identified").replace("of these findings", "of this finding")
            set_text(pe, new)
        else:
            warns.append("Executive Summary sentence about 'low-severity findings' assumes all findings are Low - rewrite it by hand.")

    # ---- reconnaissance block
    h = find_text("Server-Level Reconnaissance", "Heading2", i)
    if h is not None:
        a = h + 1
        b = a
        while b < len(els) and ptext(els[b]):
            b += 1
        proto = els[a]
        for e in els[a:b]:
            body.remove(e)
        anchor = els[h]
        if plan["recon"]:
            for t in plan["recon"]:
                np_ = copy.deepcopy(proto)
                set_text(np_, t)
                anchor.addnext(np_)
                anchor = np_
        else:
            body.remove(els[h])

    # ---- overview and scope
    def rebuild_list(label, items):
        k = find_text(label)
        a = k + 1
        b = a
        while b < len(els) and ptext(els[b]):
            b += 1
        proto = els[a]
        for e in els[a:b]:
            body.remove(e)
        anchor = els[k]
        for t in items:
            np_ = copy.deepcopy(proto)
            set_text(np_, t)
            anchor.addnext(np_)
            anchor = np_
    rebuild_list("Test Dates", plan["dates"])
    rebuild_list("Test Scope", plan["scope"])

    # ---- findings
    first_tbl = find(lambda e: e.tag == qn("w:tbl") and "Vulnerability ID" in "".join(e.itertext()))
    appendix = find_text("Appendix A", "Heading1")
    s = lambda idx: els[idx]
    tbl_p = copy.deepcopy(els[first_tbl])
    # prototypes from the first finding
    seg = els[first_tbl:appendix]
    h2 = next(e for e in seg if e.tag == qn("w:p") and pstyle(e) == "Heading2" and ptext(e).startswith("Description"))
    bodyp = els[els.index(h2) + 1]
    linkp = next(e for e in seg if e.find(qn("w:hyperlink")) is not None)
    stepp = next(e for e in seg if e.tag == qn("w:p") and e.find(".//" + qn("w:numPr")) is not None)
    imgp = next(e for e in seg if e.tag == qn("w:p") and e.find(".//" + qn("w:drawing")) is not None)
    protos = {k: copy.deepcopy(v) for k, v in dict(h2=h2, body=bodyp, link=linkp, step=stepp, img=imgp).items()}
    for e in protos.values():
        strip_bookmarks(e)
    proto_numid = int(protos["step"].find(".//" + qn("w:numId")).get(qn("w:val")))
    # remove old findings region (keep nothing between first table and appendix heading)
    for e in els[first_tbl:appendix]:
        body.remove(e)
    # remove blank paragraphs directly preceding the appendix heading
    app = els[appendix]
    new_blocks = []
    bm_id = [9000]
    img_id = [9000]

    def H(label):
        e = copy.deepcopy(protos["h2"])
        set_text(e, label + " ")
        return e

    def P(text, kind="body"):
        e = copy.deepcopy(protos[kind])
        set_text(e, text)
        return e

    def link(url):
        e = copy.deepcopy(protos["link"])
        hl = e.find(qn("w:hyperlink"))
        rid = d.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hl.set(qn("r:id"), rid)
        for t in hl.iter(qn("w:t")):
            t.text = url
        return e

    for fi, f in enumerate(plan["findings"]):
        block = []
        t = copy.deepcopy(tbl_p)
        ps = list(t.iter(qn("w:p")))
        texts = [p for p in ps]
        # row 2: left cell paragraphs [findings, title]; right cell [cvss, sev, status]
        row2 = t.findall(qn("w:tr"))[1]
        cells = row2.findall(qn("w:tc"))
        lp = cells[0].findall(qn("w:p"))
        rp = cells[1].findall(qn("w:p"))
        set_text(lp[0], f"Findings: {f['id']}")
        set_text(lp[1], f["title"] + " ")
        ppr = lp[1].find(qn("w:pPr"))
        ps_ = ppr.find(qn("w:pStyle"))
        if ps_ is not None:
            ppr.remove(ps_)
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), "1")
        ppr.insert(0, ol)
        sp_ = OxmlElement("w:spacing")
        sp_.set(qn("w:after"), "200")
        sp_.set(qn("w:lineRule"), "auto")
        ppr.insert(0, sp_)
        kn = OxmlElement("w:keepNext")
        ppr.insert(0, kn)
        for r in lp[1].findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            b = OxmlElement("w:b")
            b.set(qn("w:val"), "1")
            rpr.insert(0, b)
        bm_name = f"_vapt_finding_{fi}"
        for bs in lp[1].findall(qn("w:bookmarkStart")):
            bs.set(qn("w:name"), bm_name)
            bs.set(qn("w:id"), str(bm_id[0] + fi))
        for be in lp[1].findall(qn("w:bookmarkEnd")):
            be.set(qn("w:id"), str(bm_id[0] + fi))
        set_text(rp[0], f"CVSS: {f['cvss']}")
        set_text(rp[1], f"SEVERITY: {f['severity']} ")
        set_text(rp[2], f"STATUS: {f['status']} ")
        f["_bookmark"] = bm_name
        if fi > 0:
            block.append(spacer())
        block.append(t)
        block.append(blank())
        for label, key in (("Description:", "description"), ("Corrective Actions:", "corrective")):
            block.append(H(label))
            block += [P(x) for x in f[key]]
            block.append(blank())
        block.append(H("Reference:"))
        for x in f["reference"]:
            urls = re.findall(r"https?://\S+?(?=https?://|\s|$)", x)
            block += [link(u) for u in urls] if urls else [P(x)]
        block.append(spacer())
        block.append(H("Details of finding:"))
        block += [P(x) for x in f["details"]]
        block.append(blank())
        block.append(H("Affected Hosts:"))
        block += [P(x) for x in f["hosts"]]
        block.append(blank())
        block.append(H("Steps to Reproduce:"))
        nid = new_num(d, proto_numid)
        for x in f["steps"]:
            e = P(x, "step")
            e.find(".//" + qn("w:numId")).set(qn("w:val"), str(nid))
            block.append(e)
        block.append(blank())
        for blob in f["images"]:
            e = copy.deepcopy(protos["img"])
            rid, image = d.part.get_or_add_image(io.BytesIO(blob))
            e.find(".//" + qn("a:blip")).set(qn("r:embed"), rid)
            cx = 5731200
            cy = int(cx * image.px_height / image.px_width)
            for ext in list(e.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")) + \
                    list(e.iter(qn("a:ext"))):
                ext.set("cx", str(cx))
                ext.set("cy", str(cy))
            img_id[0] += 1
            dp = next(e.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))
            dp.set("id", str(img_id[0]))
            block.append(e)
        new_blocks += block
    for e in new_blocks:
        app.addprevious(e)

    # ---- page structure: drop blank padding, use real page breaks
    els = list(body)
    starts = []
    for e in els:
        if e.tag != qn("w:p"):
            continue
        t = ptext(e).strip()
        if (pstyle(e) == "Heading1" and t in ("Executive Summary", "Overview And Scope", "Penetration Test Results")) \
                or t.startswith("Appendix A") and pstyle(e) == "Heading1" \
                or (pstyle(e) == "Heading2" and t.startswith("Vulnerability Results")) \
                or t.startswith("©"):
            starts.append(e)
    for e in starts:
        prev = e.getprevious()
        while prev is not None and is_blank(prev):
            nxt = prev.getprevious()
            body.remove(prev)
            prev = nxt
        e.addprevious(spacer())
    # long remaining padding runs
    run = []
    toc_title = next(e for e in body if e.tag == qn("w:p") and ptext(e).strip() == "Table Of Contents")
    after = list(body)[list(body).index(toc_title):]
    for e in after + [None]:
        if e is not None and is_blank(e):
            run.append(e)
            continue
        if len(run) >= 8:
            for x in run:
                body.remove(x)
        run = []

    # ---- table of contents
    sdt = next((e for e in body if e.tag == qn("w:sdt")), None)
    if sdt is not None:
        content = sdt.find(qn("w:sdtContent"))
        sec = d.sections[0]
        text_w = int(sec.page_width.twips - sec.left_margin.twips - sec.right_margin.twips)
        for tab in sdt.iter(qn("w:tab")):
            if tab.get(qn("w:val")) == "right":
                tab.set(qn("w:pos"), str(text_w))
        entries = content.findall(qn("w:p"))
        subs = [e for e in entries if (e.find(qn("w:pPr") + "/" + qn("w:ind")) is not None
                and e.find(qn("w:pPr") + "/" + qn("w:ind")).get(qn("w:left")) == "360")]
        if subs:
            proto = copy.deepcopy(subs[0])
            anchor_prev = subs[0].getprevious()
            for e in subs:
                content.remove(e)
            for f in reversed(plan["findings"]):
                n = copy.deepcopy(proto)
                n.find(qn("w:hyperlink")).set(qn("w:anchor"), f["_bookmark"])
                ts = list(n.iter(qn("w:t")))
                ts[0].text = f["title"]
                anchor_prev.addnext(n)

    force_font(d)
    out = io.BytesIO()
    d.save(out)
    data = out.getvalue()
    if update_toc:
        data, w = fill_toc_numbers(data, plan)
        warns += w
    return data, warns


from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def force_font(d, font="Onest"):
    """Use Onest for every run/style (replaces Arial etc.)."""
    
    # 1. Set the baseline font for text that inherits its styling
    if 'Normal' in d.styles:
        d.styles['Normal'].font.name = font
        
    # 2. Force the font in the underlying XML Document Defaults
    doc_defaults = d.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr')
    if doc_defaults:
        rPr = doc_defaults[0]
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        # Clear out theme fonts that might override the explicit font name
        for attr in list(rFonts.attrib):
            del rFonts.attrib[attr]
        for k in ("ascii", "hAnsi", "cs", "eastAsia"):
            rFonts.set(qn("w:" + k), font)

    # 3. Your original code: scrub all explicit font overrides across the document
    roots = [d.element, d.styles.element]
    
    try:
        roots.append(d.part.numbering_part.element)
    except Exception:
        pass
        
    for sec in d.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            try:
                roots.append(part._element)
            except Exception:
                pass
                
    for root in roots:
        for rf in root.iter(qn("w:rFonts")):
            # Wipe existing attributes (especially theme fonts like w:asciiTheme)
            for k in list(rf.attrib):
                del rf.attrib[k]
            # Hardcode the target font
            for k in ("ascii", "hAnsi", "cs", "eastAsia"):
                rf.set(qn("w:" + k), font)


def render_pdf(docx_bytes):
    """Render with LibreOffice; returns pdf bytes or None."""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "r.docx")
        open(src, "wb").write(docx_bytes)
        try:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", td, src],
                           capture_output=True, timeout=180)
        except Exception:
            return None
        out = os.path.join(td, "r.pdf")
        return open(out, "rb").read() if os.path.exists(out) else None


def fill_toc_numbers(docx_bytes, plan):
    """Write printed page numbers (physical page - 1) into the TOC from a rendered copy."""
    warns = []
    pdf = render_pdf(docx_bytes)
    if not pdf:
        return docx_bytes, ["Could not render the document to compute table-of-contents page numbers; "
                            "right-click the TOC in Word and choose Update Field."]
    pd = fitz.open(stream=pdf, filetype="pdf")
    texts = [re.sub(r"\s+", "", p.get_text()) for p in pd]

    def page_of(title):
        for i in range(2, len(texts)):
            if re.sub(r"\s+", "", title) in texts[i]:
                return i  # 0-based physical; printed = physical (0-based) since cover is 0 and TOC is "Page 1"
        return None
    d = docx.Document(io.BytesIO(docx_bytes))
    sdt = next(e for e in d.element.body if e.tag == qn("w:sdt"))
    for p in sdt.find(qn("w:sdtContent")).findall(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        if len(ts) >= 2:
            pg = page_of(ts[0].text)
            if pg is None:
                warns.append(f"TOC: could not locate '{ts[0].text}' in the rendered document.")
            else:
                ts[-1].text = str(pg)
    out = io.BytesIO()
    d.save(out)
    return out.getvalue(), warns


def convert(pdf_bytes, template_bytes, cfg=None):
    rep = parse_original(pdf_bytes)
    plan = make_plan(rep, learn_library(template_bytes), cfg)
    data, warns = build_docx(plan, template_bytes)
    return data, plan, warns


if __name__ == "__main__":
    import sys
    pdf, tpl, out = sys.argv[1:4]
    data, plan, warns = convert(open(pdf, "rb").read(), open(tpl, "rb").read())
    open(out, "wb").write(data)
    print("wrote", out)
    for w in warns:
        print("WARN:", w)
    for r in plan["review"]:
        print(r)