import io

import docx
import pytest
from docx.shared import Inches

from core import template
from core.ingest import ingest as ingest_file
from core.export import ExportResult, build
from core.models import Block, DocumentModel, Run
from core.render import render


def sample_docx() -> bytes:
    d = docx.Document()
    d.core_properties.title = "Access Control Standard"
    d.core_properties.author = "Security Team"
    d.add_heading("Scope", 1)
    d.add_paragraph("This standard applies to every production system.")
    d.add_paragraph("Named accounts only.", style="List Bullet")
    d.add_paragraph("Rotate credentials every 90 days.", style="List Number")
    d.add_heading("Requirements", 2)
    p = d.add_paragraph("Administrative access requires ")
    p.add_run("multi-factor authentication").bold = True
    p.add_run(" for all staff.")
    table = d.add_table(rows=2, cols=3)
    for i, text in enumerate(("Guard", "Applies to", "Behaviour")):
        table.rows[0].cells[i].text = text
    for i, text in enumerate(("AuthGuard", "All routes", "Validates the JWT")):
        table.rows[1].cells[i].text = text
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def test_docx_ingest_keeps_structure():
    model = ingest_file("standard.docx", sample_docx())
    assert model.title == "Access Control Standard"
    kinds = [b.kind for b in model.blocks]
    assert "heading" in kinds and "bullet" in kinds and "number" in kinds
    assert "table" in kinds
    assert model.words() > 15
    headings = dict(model.headings())
    assert headings.get(1) == "Scope"
    assert headings.get(2) == "Requirements"


def test_render_blanks_control_number_placeholder():
    model = ingest_file("standard.docx", sample_docx())
    spec = template.get("portrait")
    data, _ = render(model, spec.read(), cover=True)
    out = docx.Document(io.BytesIO(data))

    headers = [p.text for s in out.sections for p in
               (s.header.paragraphs + s.first_page_header.paragraphs)]
    # walk textboxes too
    from docx.oxml.ns import qn
    all_text = []
    for s in out.sections:
        for part in (s.header, s.first_page_header, s.footer, s.first_page_footer):
            for p in part._element.iter(qn("w:p")):
                all_text.append("".join(t.text or "" for t in p.iter(qn("w:t"))))
    assert not any("CONTROL#" in t for t in all_text)


def test_render_fits_tables_to_the_page():
    model = inject_wide_table()
    spec = template.get("landscape")
    data, _ = render(model, spec.read())
    out = docx.Document(io.BytesIO(data))
    sec = out.sections[0]
    from docx.shared import Emu
    usable = int(Emu(sec.page_width - sec.left_margin - sec.right_margin).twips)
    from docx.oxml.ns import qn
    tbl = out.tables[0]._tbl
    width = int(tbl.find(qn("w:tblPr")).find(qn("w:tblW")).get(qn("w:w")))
    assert abs(width - usable) < 5
    assert len(out.tables[0].columns) == 6
    assert len(out.tables[0].rows) == 4


def inject_wide_table() -> DocumentModel:
    rows = [["A", "B", "C", "D", "E", "F"]] + [["1"] * 6] * 3
    return DocumentModel(meta={"title": "Wide"}, blocks=[Block(kind="table", rows=rows)])


def test_heading_levels_follow_visual_prominence():
    model = DocumentModel(meta={"title": "Doc"})
    for level, text in ((1, "Section"), (2, "Subsection"), (3, "Detail")):
        block = Block(kind="heading", level=level)
        block.set_text(text)
        model.blocks.append(block)
    data, _ = render(model, template.get("portrait").read())
    out = docx.Document(io.BytesIO(data))
    by_text = {p.text: p.style.style_id for p in out.paragraphs}
    # template Heading1 is 11pt; Heading2 is 16pt, Heading3 14pt, Heading4 12pt
    assert by_text["Section"] == "Heading2"
    assert by_text["Subsection"] == "Heading3"
    assert by_text["Detail"] == "Heading4"


def test_no_cover_uses_the_normal_first_page_header():
    model = DocumentModel(meta={"title": "No cover"})
    model.blocks = [Block(kind="paragraph", runs=[Run(text="Body")])]
    data, _ = render(model, template.get("portrait").read(), cover=False)
    out = docx.Document(io.BytesIO(data))
    assert out.sections[0].different_first_page_header_footer is False
    assert out.paragraphs[0].style.name == "Title"


def test_lists_have_nested_levels_and_restart_numbering():
    from docx.oxml.ns import qn

    model = DocumentModel(meta={"title": "Lists"})
    items = [("number", 0, "First"), ("number", 0, "Second"),
             ("paragraph", 0, "Break"), ("number", 0, "Again"),
             ("bullet", 0, "Bullet"), ("bullet", 1, "Nested")]
    for kind, level, text in items:
        block = Block(kind=kind, level=level)
        block.set_text(text)
        model.blocks.append(block)

    data, _ = render(model, template.get("portrait").read())
    out = docx.Document(io.BytesIO(data))
    by_text = {p.text: p for p in out.paragraphs}

    def num_id(text):
        node = by_text[text]._p.find(
            qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
        return node.get(qn("w:val"))

    assert num_id("First") == num_id("Second")
    assert num_id("First") != num_id("Again")

    ilvl = by_text["Nested"]._p.find(
        qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:ilvl"))
    assert ilvl.get(qn("w:val")) == "1"

    numbering = out.part.numbering_part.element
    first_abstract = numbering.find(qn("w:abstractNum"))
    assert len(first_abstract.findall(qn("w:lvl"))) == 5


def test_auto_orientation():
    narrow = ingest_file("standard.docx", sample_docx())
    assert template.choose(narrow, "auto")[0] == "portrait"
    wide = inject_wide_table()
    assert template.choose(wide, "auto")[0] == "landscape"


def test_fidelity_by_word_count():
    model = ingest_file("standard.docx", sample_docx())
    data, _ = render(model, template.get("portrait").read())
    out = docx.Document(io.BytesIO(data))
    text = " ".join(p.text for p in out.paragraphs)
    for table in out.tables:
        for row in table.rows:
            for cell in row.cells:
                text += " " + cell.text
    out_words = len(text.split())
    assert out_words >= model.words() * 0.9, (out_words, model.words())


@pytest.mark.parametrize("name,payload,expected", [
    ("notes.txt", b"Work Instruction\n\nDo the thing.\n\nThen sign it off.\n", "paragraph"),
    ("readme.md", b"# Title\n\nSome text.\n\n- one\n- two\n\n| a | b |\n|---|---|\n| 1 | 2 |\n", "table"),
    ("page.html", b"<html><body><h1>Guide</h1><p>Hello</p><ul><li>x</li></ul>"
                  b"<table><tr><th>A</th></tr><tr><td>1</td></tr></table></body></html>", "table"),
    ("rows.csv", b"name,role\nAna,Analyst\nBo,Engineer\n", "table"),
])
def test_other_ingest_paths(name, payload, expected):
    model = ingest_file(name, payload)
    assert model.blocks
    assert any(b.kind == expected for b in model.blocks)


def test_txt_heading_detection():
    model = ingest_file("notes.txt", b"Work Instruction\n\nDo the thing.\n")
    assert model.meta["title"] == "Work Instruction"

