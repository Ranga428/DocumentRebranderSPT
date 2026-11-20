import io

import docx
from PIL import Image

from core.ingest import ingest as ingest_file
from core.render import render
from core import template


def png_bytes() -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", (200, 100), (46, 125, 50)).save(buf, format="PNG")
    return buf.getvalue()


def test_images_survive_ingest_and_render():
    d = docx.Document()
    d.add_paragraph("Illustrated guide")
    d.add_paragraph("Before:")
    d.add_picture(io.BytesIO(png_bytes()))
    d.add_paragraph("After:")
    buf = io.BytesIO()
    d.save(buf)

    model = ingest_file("guide.docx", buf.getvalue())
    assert any(b.kind == "image" for b in model.blocks)

    data, _ = render(model, template.get("portrait").read())
    out = docx.Document(io.BytesIO(data))
    blips = out.element.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    assert len(blips) >= 1
