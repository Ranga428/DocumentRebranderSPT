"""The Streamlit workspace accepts several uploads at once."""
import io
from pathlib import Path

import docx
import pytest

pytest.importorskip("streamlit.testing.v1")
from streamlit.testing.v1 import AppTest  # noqa: E402

APP = Path(__file__).resolve().parents[1] / "app.py"
MIME = ("application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document")


def _docx_bytes(title: str) -> bytes:
    d = docx.Document()
    d.add_heading(title, 1)
    d.add_paragraph(f"Body of {title}.")
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def test_uploads_are_added_to_the_workspace_and_can_be_removed():
    at = AppTest.from_file(str(APP), default_timeout=120).run()
    uploader = at.get("file_uploader")[0]
    uploader.set_value([
        ("one.docx", _docx_bytes("One"), MIME),
        ("two.docx", _docx_bytes("Two"), MIME),
    ]).run()

    assert not list(at.exception)
    assert len(at.session_state["order"]) == 2
    assert at.selectbox[0].options == ["01 · one.docx — One", "02 · two.docx — Two"]

    at.button[0].click().run()  # the first workspace "✕"
    assert not list(at.exception)
    assert len(at.session_state["order"]) == 1

    # the file is still in the uploader widget, but must stay removed
    at.run()
    assert len(at.session_state["order"]) == 1
    assert len(at.session_state["names"]) == 1
