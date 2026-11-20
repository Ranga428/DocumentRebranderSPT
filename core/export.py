"""DOCX/PDF export with table-of-contents page numbering."""
from __future__ import annotations

import io
from dataclasses import dataclass, field

import docx
import pymupdf as fitz
from docx.oxml.ns import qn

from . import office
from .models import DocumentModel
from .render import render


@dataclass
class ExportResult:
    docx: bytes
    pdf: bytes | None = None
    warnings: list[str] = field(default_factory=list)
    pages: int = 0


def render_pdf(docx_bytes: bytes, timeout: int = 180) -> tuple[bytes | None, str | None]:
    """Convert to PDF, returning (pdf, error)."""
    if not office.available():
        return None, "LibreOffice was not found"
    try:
        return office.to_pdf(docx_bytes, "rebrand.docx", timeout=timeout), None
    except Exception as exc:
        return None, str(exc)


def build(model: DocumentModel, template_bytes: bytes, *, cover: bool = True,
          control_number: str = "", timeout: int = 180) -> ExportResult:
    warnings: list[str] = []
    docx_bytes, render_warnings = render(
        model, template_bytes, cover=cover, control_number=control_number)
    warnings += render_warnings

    pdf, error = render_pdf(docx_bytes, timeout=timeout)
    if pdf is None:
        warnings.append(
            f"PDF export is unavailable: {error or 'conversion failed'}. "
            "The .docx download still works.")
    pages = 0
    if pdf:
        doc = fitz.open(stream=pdf, filetype="pdf")
        pages = doc.page_count
        doc.close()
    warnings = list(dict.fromkeys(warnings))
    return ExportResult(docx=docx_bytes, pdf=pdf, warnings=warnings, pages=pages)


def thumbnails(pdf: bytes, limit: int = 12, dpi: int = 60) -> list[bytes]:
    out = []
    doc = fitz.open(stream=pdf, filetype="pdf")
    for i, page in enumerate(doc):
        if i >= limit:
            break
        out.append(page.get_pixmap(dpi=dpi).tobytes("png"))
    doc.close()
    return out
