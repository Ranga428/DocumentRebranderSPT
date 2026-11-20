"""Paragraph grouping: bullets must not swallow the sections that follow."""
import io

import docx

from core import template
from core.ingest import _heading_segments, ingest
from core.models import Block, DocumentModel, Run
from core.render import render


def docx_bytes(*paragraphs: str) -> bytes:
    d = docx.Document()
    for text in paragraphs:
        d.add_paragraph(text)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def test_bullet_stops_before_numbered_heading_and_body():
    model = ingest("case.docx", docx_bytes(
        "\u2022 Shared code: the shared core library must be built before the apps.",
        "4. Core licensing module",
        "LicensingModule imports the RBAC service and enforces the 10 MB file-size limit.",
        "4.1 Controllers",
        "Controllers handle HTTP requests and return JSON.",
    ))
    shape = [(b.kind, b.level, b.text[:35]) for b in model.blocks]
    assert [k for k, _, _ in shape] == ["bullet", "heading", "paragraph", "heading", "paragraph"]
    assert shape[1][1] == 2 and shape[1][2] == "4. Core licensing module"
    assert shape[3][1] == 3 and shape[3][2] == "4.1 Controllers"
    # the heading and the body are separate paragraphs, not glued to the bullet
    assert model.blocks[0].text.endswith("the apps.")
    assert model.blocks[2].text.startswith("LicensingModule")


def test_heading_glued_after_sentence_is_split_out():
    model = ingest("glued.docx", docx_bytes(
        "LicensingModule imports the RBAC service and enforces a 10 MB file-size limit. 4.1 Controllers",
    ))
    assert [b.kind for b in model.blocks] == ["paragraph", "heading"]
    assert model.blocks[1].level == 3
    assert model.blocks[1].text == "4.1 Controllers"
    assert all(r.bold for r in model.blocks[1].runs if r.text.strip())


def test_bullet_keeps_a_lowercase_wrapped_continuation():
    model = ingest("wrap.docx", docx_bytes(
        "\u2022 Shared code: built before the apps",
        "and reused by every module.",
    ))
    assert len(model.blocks) == 1
    assert model.blocks[0].kind == "bullet"
    assert model.blocks[0].text.endswith("and reused by every module.")


def test_quantities_are_not_treated_as_headings():
    text = "The uploader accepts a 10 MB file-size limit and a 12.5 mm margin."
    assert _heading_segments(text) == [("text", text)]


def test_numbered_heading_renders_bold():
    model = DocumentModel(meta={"title": "Doc"})
    model.blocks = [
        Block(kind="paragraph", runs=[Run(text="Intro.")]),
        Block(kind="heading", level=2, runs=[Run(text="4. Core licensing module", bold=True)]),
    ]
    data, _ = render(model, template.get("portrait").read())
    out = docx.Document(io.BytesIO(data))
    heading = next(p for p in out.paragraphs if p.text == "4. Core licensing module")
    assert heading.style.style_id.startswith("Heading")
    assert all(r.bold for r in heading.runs if r.text.strip())


def test_bullet_lead_in_keeps_its_space_after_the_colon():
    model = DocumentModel(meta={"title": "Doc"})
    model.blocks = [
        Block(kind="bullet", runs=[Run(text="Shared code: the core is built first.")]),
    ]
    data, _ = render(model, template.get("portrait").read(), cover=False)
    out = docx.Document(io.BytesIO(data))
    bullet = next(p for p in out.paragraphs if "Shared code" in p.text)
    assert bullet.runs[0].text == "Shared code:"
    assert bullet.runs[0].bold is True
    assert bullet.runs[1].text.startswith(" ")

