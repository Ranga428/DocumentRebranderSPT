import io

import pymupdf as fitz
import pytest

from core import office, template
from core.export import thumbnails
from core.ingest import ingest as ingest_file
from core.render import render


def sample_pdf() -> bytes:
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 90), "Access Control Policy", fontsize=22)
    page.insert_text((72, 130), "Scope", fontsize=16)
    page.insert_text((72, 155), "This policy applies to all production systems "
                                "and every member of staff.", fontsize=11)
    page.insert_text((72, 175), "Second paragraph of body text for the document.",
                     fontsize=11)
    page2 = doc.new_page()
    page2.insert_text((72, 90), "Requirements", fontsize=16)
    page2.insert_text((72, 115), "Use named accounts. Rotate credentials regularly.",
                      fontsize=11)
    return doc.tobytes()


def test_pdf_headings_and_body_are_detected():
    model = ingest_file("policy.pdf", sample_pdf())
    assert model.title == "Access Control Policy"
    headings = [b for b in model.blocks if b.kind == "heading"]
    assert [h.text for h in headings][:1] == ["Scope"]
    assert any("production systems" in b.text for b in model.blocks)


def test_pdf_render_keeps_text():
    import docx as docx_module

    model = ingest_file("policy.pdf", sample_pdf())
    data, _ = render(model, template.get("portrait").read())
    out = docx_module.Document(io.BytesIO(data))
    text = " ".join(p.text for p in out.paragraphs)
    assert "Access Control Policy" in text
    assert "production systems" in text
    assert len(model.blocks) >= 4


def test_pdf_page_thumbnails():
    pdf = sample_pdf()
    pages = thumbnails(pdf, limit=4)
    assert len(pages) == 2
    assert pages[0][:4] == b"\x89PNG"
