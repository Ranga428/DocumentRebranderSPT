"""Document Rebrander — Streamlit front end.

Pipeline: any supported document -> structured model -> brand template -> DOCX/PDF.
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

import streamlit as st
from core import office, template
import converter
from core.export import build as build_export
from core.export import thumbnails
from core.ingest import IngestError, accept_types, ingest
from core.models import Block, DocumentModel, Run
from core.render import apply_control_number_bytes

ROOT = Path(__file__).resolve().parent
LOGO = ROOT / "SapientLogoNoText.png"

st.set_page_config(page_title="Document Rebrander", page_icon=str(LOGO), layout="wide")

ORIENTATIONS = {"auto": "Auto (fit content)", "portrait": "Portrait", "landscape": "Landscape"}


def _logo_html() -> str:
    if not LOGO.exists():
        return ""
    data = base64.b64encode(LOGO.read_bytes()).decode("ascii")
    return (f'<img src="data:image/png;base64,{data}" alt="Sapient logo" '
            'style="display:block;width:48px;height:48px;">')


def _slug(text: str, fallback: str = "document") -> str:
    text = re.sub(r"[^\w\s-]", "", (text or "").strip()).strip()
    return re.sub(r"[\s_-]+", "-", text).lower()[:60] or fallback


@st.cache_data(show_spinner=False)
def _spec_report(key: str, mtime: float) -> tuple[bytes, list[str]]:
    """Cached template bytes + contract validation (invalidated by mtime)."""
    data = template.get(key).read()
    return data, template.validate(data, key)


@st.cache_data(show_spinner=False, max_entries=4)
def _thumbs(pdf: bytes, limit: int = 9) -> list[bytes]:
    return thumbnails(pdf, limit=limit)


def _signature(model: DocumentModel, options: dict, template_hash: str) -> str:
    payload = {
        "meta": model.meta,
        "blocks": [
            {"kind": b.kind, "level": b.level, "text": b.text, "rows": b.rows,
             "header": b.header_row,
             "image": hashlib.sha256(b.image).hexdigest() if b.image else None}
            for b in model.blocks
        ],
        "options": options,
        "template": template_hash,
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True, default=str).encode()).hexdigest()


# --------------------------------------------------------------------- state
def _init() -> None:
    defaults = {
        "docs": {},            # digest -> DocumentModel
        "names": {},           # digest -> uploaded filename
        "order": [],           # digests, in upload order
        "active": None,        # digest currently being edited
        "failed": {},          # digest -> error message (do not retry)
        "results": {},         # digest -> ExportResult
        "result_sigs": {},     # digest -> edit signature
        "result_keys": {},     # digest -> template key
        "build_errors": {},    # digest -> build error message
        "dismissed": set(),    # upload ids removed from the workspace
        "uploader_epoch": 0,   # bump to empty the uploader widget
        "vapt_results": [],    # list of converted VAPT reports
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _reset_editor() -> None:
    # only widget state owned by the editor; sidebar settings stay as chosen
    for key in list(st.session_state.keys()):
        if key.startswith(("txt_", "lvl_", "tbl_", "hdr_", "img_", "meta_")):
            del st.session_state[key]


def _active_model() -> DocumentModel | None:
    active = st.session_state.active
    return st.session_state.docs.get(active) if active else None


def _active_name() -> str:
    return st.session_state.names.get(st.session_state.active or "", "document")


def _display_name(digest: str) -> str:
    name = st.session_state.names.get(digest, digest[:8])
    model = st.session_state.docs.get(digest)
    title = (model.title or "").strip() if model else ""
    return f"{name} — {title[:60]}" if title and title not in name else name


def _drop_document(digest: str) -> None:
    st.session_state.dismissed.add(digest)
    st.session_state.docs.pop(digest, None)
    st.session_state.names.pop(digest, None)
    st.session_state.failed.pop(digest, None)
    st.session_state.results.pop(digest, None)
    st.session_state.result_sigs.pop(digest, None)
    st.session_state.result_keys.pop(digest, None)
    st.session_state.build_errors.pop(digest, None)
    if digest in st.session_state.order:
        st.session_state.order.remove(digest)
    if st.session_state.active == digest:
        st.session_state.active = (st.session_state.order[0]
                                   if st.session_state.order else None)


def _base_name(digest: str) -> str:
    model = st.session_state.docs.get(digest)
    title = (model.title if model else "") or Path(st.session_state.names.get(digest, "")).stem
    return _slug(title)


def _zip_downloads(digests: list[str]) -> bytes:
    buf = io.BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for digest in digests:
            result = st.session_state.results.get(digest)
            if result is None:
                continue
            base = _base_name(digest)
            if base in used:
                base = f"{base}-{digest[:6]}"
            used.add(base)
            zf.writestr(f"{base}-SPT.docx", result.docx)
            if result.pdf:
                zf.writestr(f"{base}-SPT.pdf", result.pdf)
    return buf.getvalue()


def _zip_vapt(items: list[dict]) -> bytes:
    buf = io.BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in items:
            if not item.get("docx"):
                continue
            base = _slug(Path(item["name"]).stem, "vapt-report")
            if base in used:
                base = f"{base}-{hashlib.sha256(item['name'].encode()).hexdigest()[:6]}"
            used.add(base)
            zf.writestr(f"{base}-SPT.docx", item["docx"])
            if item.get("pdf"):
                zf.writestr(f"{base}-SPT.pdf", item["pdf"])
    return buf.getvalue()


def _build_documents(digests: list[str], orientation: str, cover: bool) -> None:
    """Build every requested document, storing results in session state."""
    progress = st.progress(0.0, text="Building…")
    errors: list[str] = []
    for done, digest in enumerate(digests, start=1):
        model = st.session_state.docs.get(digest)
        if model is None:
            continue
        name = st.session_state.names.get(digest, digest[:8])
        try:
            key, _warning = template.choose(model, orientation)
            spec = template.get(key)
            data = spec.read()
            result = build_export(model, data, cover=cover)
        except Exception as exc:  # pragma: no cover - defensive
            errors.append(f"{name}: {exc}")
            st.session_state.build_errors[digest] = str(exc)
            continue
        st.session_state.results[digest] = result
        st.session_state.result_keys[digest] = key
        st.session_state.result_sigs[digest] = _signature(
            model, {"orientation": key, "cover": cover}, spec.sha256)
        st.session_state.build_errors.pop(digest, None)
        progress.progress(done / len(digests), text=f"Built {name}")
    progress.empty()
    for error in errors:
        st.error(f"Build failed — {error}")



# ------------------------------------------------------------- block actions
def _index(model: DocumentModel, block_id: str) -> int | None:
    for i, block in enumerate(model.blocks):
        if block.id == block_id:
            return i
    return None


def _move(block_id: str, delta: int) -> None:
    model = _active_model()
    if model is None:
        return
    i = _index(model, block_id)
    if i is None:
        return
    j = max(0, min(len(model.blocks) - 1, i + delta))
    if i != j:
        model.blocks[i], model.blocks[j] = model.blocks[j], model.blocks[i]


def _delete(block_id: str) -> None:
    model = _active_model()
    if model is None:
        return
    i = _index(model, block_id)
    if i is not None:
        model.blocks.pop(i)


def _duplicate(block_id: str) -> None:
    import copy
    import uuid
    model = _active_model()
    if model is None:
        return
    i = _index(model, block_id)
    if i is not None:
        clone = copy.deepcopy(model.blocks[i])
        clone.id = uuid.uuid4().hex[:10]
        model.blocks.insert(i + 1, clone)


_init()

# --------------------------------------------------------------------- header
st.markdown(
    f"""
    <div style="display:flex;align-items:center;gap:1rem;">
        {_logo_html()}
        <div>
            <h1 style="margin:0;font-size:2.4rem;line-height:1.1;">Document Rebrander</h1>
            <div style="opacity:.7;margin-top:.2rem;">
                Any document in, SPT-branded Word or PDF out.
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.write("")

# apply a pending cover-page default (set after an upload) before the widget exists
if "_cover_pending" in st.session_state:
    st.session_state["cover_choice"] = st.session_state.pop("_cover_pending")

# -------------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Brand settings")

    orientation = st.radio(
        "Page orientation", list(ORIENTATIONS), format_func=ORIENTATIONS.get,
        key="orient_choice")
    cover = st.checkbox("Add a cover page",
        value=st.session_state.get("cover_choice", False), key="cover_choice")

    st.divider()
    st.subheader("Template status")
    for key, spec in template.available().items():
        if not spec.path.exists():
            st.warning(f"{key.title()}: file missing")
            continue
        issues = _spec_report(key, spec.path.stat().st_mtime)[1]
        if issues:
            st.warning(f"{key.title()}: {issues[0]}")
            with st.expander(f"{len(issues)} template issue(s)"):
                for issue in issues:
                    st.write(f"- {issue}")
        else:
            st.success(f"{key.title()} — ready")
    st.caption(f"Font: {template.manifest().get('font', 'Onest')} · "
               f"manifest v{template.manifest().get('version', '?')}")

    st.divider()
    if office.available():
        st.success("LibreOffice detected")
    else:
        st.warning("LibreOffice not found — PDF export and TOC page numbers are "
                   "disabled. DOCX still works.")

# -------------------------------------------------------------------- mode
mode = st.radio("Mode", ["General document", "VAPT report"], horizontal=True,
                key="mode_choice", help="General: rebrand any document. VAPT: convert a "
                "VAPT assessment PDF into the branded report format.")

if mode == "VAPT report":
    st.subheader("VAPT report conversion")
    st.caption("VAPT mode uses the legacy converter and requires a VAPT-specific "
               "template containing a findings library. The generic brand templates "
               "will not work for VAPT conversion. Select several PDFs to convert "
               "them in one batch.")
    vapt_uploads = st.file_uploader("VAPT assessment PDFs", type=["pdf"],
                                    accept_multiple_files=True, key="vapt_upload")
    import os as _os
    vapt_dir = _os.path.join(str(ROOT), "vapt_templates")
    vapt_files = sorted(f for f in _os.listdir(vapt_dir)
                        if f.lower().endswith(".docx")) if _os.path.isdir(vapt_dir) else []
    if not vapt_files:
        st.error("No VAPT templates found in vapt_templates/.")
        st.stop()
    vapt_template_name = st.selectbox("VAPT template", vapt_files,
                                      key="vapt_template_name")
    count = len(vapt_uploads or [])
    convert_label = f"Convert {count} VAPT report" + ("" if count == 1 else "s")
    if st.button(convert_label, type="primary", disabled=not vapt_uploads):
        with open(_os.path.join(vapt_dir, vapt_template_name), "rb") as fh:
            tpl_bytes = fh.read()
        items: list[dict] = []
        progress = st.progress(0.0, text="Converting…")
        for done, upload in enumerate(vapt_uploads, start=1):
            item = {"name": upload.name, "docx": None, "pdf": None,
                    "plan": {}, "warnings": [], "error": None}
            try:
                data, plan, warns = converter.convert(upload.getvalue(), tpl_bytes)
                item["docx"] = data
                item["plan"] = plan
                item["warnings"] = warns
                item["pdf"] = converter.render_pdf(data)
            except Exception as exc:
                item["error"] = (str(exc) + "  (VAPT mode needs a VAPT-specific "
                                 "template containing a findings library; the generic "
                                 "brand templates will not work.)")
            items.append(item)
            progress.progress(done / len(vapt_uploads), text=f"Converted {upload.name}")
        progress.empty()
        st.session_state["vapt_results"] = items

    items = st.session_state.get("vapt_results", [])
    if items:
        ok = [item for item in items if item.get("docx")]
        if len(ok) > 1:
            st.download_button(
                f"⬇ Download all {len(ok)} reports (.zip)", _zip_vapt(ok),
                "vapt-reports-SPT.zip", "application/zip", use_container_width=True)
        for n, item in enumerate(items, start=1):
            base = _slug(Path(item["name"]).stem, "vapt-report")
            label = f"{n:02d} · {item['name']}"
            if item.get("error"):
                label += " — failed"
            with st.expander(label, expanded=len(items) == 1):
                if item.get("error"):
                    st.error("Conversion failed: " + item["error"])
                    continue
                warns = item.get("warnings", [])
                if warns:
                    with st.expander(str(len(warns)) + " conversion warning(s)"):
                        for w in warns:
                            st.warning(w)
                plan = item.get("plan", {})
                review = plan.get("review", []) if isinstance(plan, dict) else []
                if review:
                    with st.expander("Review items (" + str(len(review)) + ")"):
                        for title, field, src in review[:60]:
                            st.write("- **" + str(title) + "** - " + str(field) + ": " + str(src))
                d1, d2 = st.columns(2)
                d1.download_button("⬇ Download .docx", item["docx"], base + "-SPT.docx",
                                   "application/vnd.openxmlformats-officedocument."
                                   "wordprocessingml.document",
                                   use_container_width=True, key=f"vapt_docx_{n}")
                if item.get("pdf"):
                    d2.download_button("⬇ Download .pdf", item["pdf"], base + "-SPT.pdf",
                                       "application/pdf", use_container_width=True,
                                       key=f"vapt_pdf_{n}")
                else:
                    d2.info("PDF export unavailable (LibreOffice not found).")
                if item.get("pdf"):
                    previews = _thumbs(item["pdf"], 3)
                    cols = st.columns(3)
                    for off, img in enumerate(previews):
                        cols[off].image(img, caption="Page " + str(off + 1),
                                        use_container_width=True)
    st.stop()

# -------------------------------------------------------------------- upload
uploads = st.file_uploader(
    "Source documents",
    type=accept_types(),
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_epoch}",
    help="PDF, Word, OpenDocument, RTF, HTML, Markdown, text, CSV, Excel or PowerPoint. "
         "Select several files at once to build them as a batch.",
)

if uploads:
    # ids still sitting in the uploader; a document removed from the workspace
    # must not be re-ingested until its file is removed from the uploader too
    present = {upload.file_id for upload in uploads if upload.file_id}
    st.session_state.dismissed.intersection_update(present)
    added: str | None = None
    for upload in uploads:
        key = upload.file_id or hashlib.sha256(upload.getvalue()).hexdigest()
        if (key in st.session_state.docs or key in st.session_state.failed
                or key in st.session_state.dismissed):
            continue
        data = upload.getvalue()
        with st.spinner(f"Reading {upload.name}…"):
            try:
                model = ingest(upload.name, data)
            except IngestError as exc:
                st.session_state.failed[key] = f"{upload.name}: {exc}"
                continue
            except Exception as exc:  # pragma: no cover - defensive
                st.session_state.failed[key] = (
                    f"{upload.name}: unexpected error while reading the file ({exc})")
                continue
        st.session_state.docs[key] = model
        st.session_state.names[key] = upload.name
        st.session_state.order.append(key)
        added = key
    for message in st.session_state.failed.values():
        st.error(message)
    if added is not None:
        st.session_state.active = added
        st.session_state["_cover_pending"] = bool(
            st.session_state.docs[added].meta.get("has_cover", False))
        _reset_editor()
        st.rerun()

order: list[str] = st.session_state.order
if not order:
    st.subheader("How it works")
    st.markdown(
        "1. Upload one or more documents — text, headings, lists, tables and images "
        "are extracted into editable outlines.\n"
        "2. Pick page orientation and output options in the sidebar.\n"
        "3. Review and edit each outline, then build the branded .docx and PDF — "
        "individually or all at once as a ZIP."
    )
    st.info("Supported: " + ", ".join(f"`.{t}`" for t in sorted(accept_types())))
    st.stop()
    raise SystemExit(0)  # st.stop() is a no-op in bare mode

# --------------------------------------------------------------- workspace
if st.session_state.active not in order:
    st.session_state.active = order[0]

if len(order) > 0:
    pick, drop = st.columns([4, 1])
    labels = {digest: f"{i + 1:02d} · {_display_name(digest)}"
              for i, digest in enumerate(order)}
    chosen = pick.selectbox(
        f"Editing document (of {len(order)})", list(labels.values()),
        index=order.index(st.session_state.active))
    st.session_state.active = next(d for d, label in labels.items() if label == chosen)
    with drop.popover("Workspace", use_container_width=True):
        st.caption(f"{len(order)} document(s) loaded")
        for digest in list(order):
            row = st.columns([4, 1])
            row[0].write(st.session_state.names.get(digest, digest[:8]))
            row[1].button("✕", key=f"rm_{digest}", help="Remove from the workspace",
                          on_click=_drop_document, args=(digest,))
        if st.button("Remove all", key="rm_all"):
            for digest in list(order):
                _drop_document(digest)
            st.session_state.dismissed.clear()
            st.session_state.uploader_epoch += 1
            st.session_state.active = None
            _reset_editor()
            st.rerun()

model: DocumentModel = st.session_state.docs[st.session_state.active]

# ------------------------------------------------------------------- summary
counts = model.counts()
cols = st.columns(4)
cols[0].metric("Blocks", len(model.blocks))
cols[1].metric("Words", f"{model.words():,}")
cols[2].metric("Tables", counts.get("table", 0))
cols[3].metric("Images", counts.get("image", 0))

if model.warnings:
    with st.expander(f"⚠️ {len(model.warnings)} import warning(s)", expanded=True):
        for warning in model.warnings:
            st.warning(warning)

st.subheader("Document details")
_mk = st.session_state.active
m1, m2, m3 = st.columns(3)
model.title = m1.text_input("Title", value=model.title, key=f"meta_title_{_mk}")
model.subtitle = m2.text_input("Subtitle", value=model.subtitle, key=f"meta_subtitle_{_mk}")
model.meta["author"] = m3.text_input("Author", value=str(model.meta.get("author", "")),
                                     key=f"meta_author_{_mk}")
m4, m5 = st.columns(2)
model.meta["date"] = m4.text_input("Date", value=str(model.meta.get("date", "")),
                                   key=f"meta_date_{_mk}")
control_value = ""

# --------------------------------------------------------------------- editor
tab_content, tab_build = st.tabs(["Content", "Build & preview"])

with tab_content:
    add1, add2, add3 = st.columns([1, 1, 3])
    if add1.button("+ Paragraph"):
        model.blocks.append(Block(kind="paragraph", runs=[Run(text="New paragraph")]))
    if add2.button("+ Heading"):
        model.blocks.append(Block(kind="heading", level=2, runs=[Run(text="New heading")]))
    if add3.button("Clean blank blocks"):
        model.blocks = [b for b in model.blocks if b.kind != "paragraph" or b.text.strip()]

    st.caption(f"{len(model.blocks)} blocks — open a block to edit it. "
               "Untouched blocks keep their original bold/italic formatting.")

    for i, block in enumerate(list(model.blocks)):
        preview = block.text.strip().replace("\n", " ")
        if block.kind == "table":
            preview = f"{len(block.rows)}×{len(block.rows[0]) if block.rows else 0} table"
        elif block.kind == "image":
            preview = "image"
        label = f"{i + 1:02d} · {block.kind.title()}" + (f" — {preview[:70]}" if preview else "")
        with st.expander(label):
            if block.kind in ("paragraph", "heading", "title", "subtitle", "quote",
                              "caption", "bullet", "number"):
                text = st.text_area("Text", value=block.text, key=f"txt_{block.id}",
                                    height=110)
                if text != block.text:
                    block.set_text(text)

            if block.kind == "heading":
                block.level = st.selectbox(
                    "Heading level", [1, 2, 3, 4, 5, 6],
                    index=max(0, min(5, (block.level or 1) - 1)),
                    key=f"lvl_{block.id}")

            if block.kind == "table":
                raw = "\n".join(" | ".join(row) for row in (block.rows or [[""]]))
                edited = st.text_area(
                    "Rows — one line per row, cells separated by |",
                    value=raw, key=f"tbl_{block.id}", height=160)
                if edited != raw:
                    block.rows = [[cell.strip() for cell in line.split("|")]
                                  for line in edited.splitlines() if line.strip()]
                block.header_row = st.checkbox("First row is a header",
                                               value=block.header_row,
                                               key=f"hdr_{block.id}")

            if block.kind == "image":
                if block.image:
                    st.image(block.image, use_container_width=True)
                replace = st.file_uploader("Replace image", type=["png", "jpg", "jpeg"],
                                           key=f"img_{block.id}")
                if replace is not None:
                    block.image = replace.getvalue()
                    block.width_px = None

            b1, b2, b3, b4 = st.columns(4)
            b1.button("↑ Up", key=f"up_{block.id}", on_click=_move,
                      args=(block.id, -1), disabled=i == 0, use_container_width=True)
            b2.button("↓ Down", key=f"down_{block.id}", on_click=_move,
                      args=(block.id, 1), disabled=i == len(model.blocks) - 1,
                      use_container_width=True)
            b3.button("Duplicate", key=f"dup_{block.id}", on_click=_duplicate,
                      args=(block.id,), use_container_width=True)
            b4.button("Delete", key=f"del_{block.id}", on_click=_delete,
                      args=(block.id,), use_container_width=True)

with tab_build:
    try:
        key, orientation_warning = template.choose(model, orientation)
        spec = template.get(key)
        template_data, issues = _spec_report(key, spec.path.stat().st_mtime)
    except template.TemplateError as exc:
        st.error(str(exc))
        st.stop()
        raise SystemExit(0)  # st.stop() is a no-op in bare mode
    if orientation_warning:
        st.warning(orientation_warning)

    c1, c2 = st.columns(2)
    c1.metric("Template", key.title())
    c2.metric("Cover page", "Yes" if cover else "No")

    if issues:
        st.error("The template has contract problems; fix these before building:")
        for issue in issues:
            st.write(f"- {issue}")

    active = st.session_state.active
    options = {"orientation": key, "cover": cover}
    signature = _signature(model, options, spec.sha256)

    if len(order) > 1:
        build_col, all_col, _filler = st.columns([1, 1, 2])
        build_all = all_col.button(f"Build all {len(order)} documents",
                                   disabled=bool(issues), use_container_width=True)
    else:
        build_col = st.columns([1, 3])[0]
        build_all = False

    if build_col.button("Build brand document", type="primary",
                        disabled=bool(issues), use_container_width=True):
        with st.spinner("Building the branded document…"):
            try:
                result = build_export(model, template_data, cover=cover)
                st.session_state.results[active] = result
                st.session_state.result_sigs[active] = signature
                st.session_state.result_keys[active] = key
                st.session_state.build_errors.pop(active, None)
            except Exception as exc:  # pragma: no cover - defensive
                st.session_state.results.pop(active, None)
                st.session_state.build_errors[active] = str(exc)

    if build_all:
        _build_documents(order, orientation, cover)

    error = st.session_state.build_errors.get(active)
    if error:
        st.error(f"Build failed: {error}")

    result = st.session_state.results.get(active)
    if result is not None:
        if st.session_state.result_sigs.get(active) != signature:
            st.info("The document changed since this build. Build again to refresh "
                    "the downloads.")
        base = _base_name(active)
        d1, d2 = st.columns(2)
        d1.download_button(
            "⬇ Download .docx", result.docx, f"{base}-SPT.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True, key="dl_docx_active")
        if result.pdf:
            d2.download_button("⬇ Download .pdf", result.pdf, f"{base}-SPT.pdf",
                               "application/pdf", use_container_width=True,
                               key="dl_pdf_active")

        if result.warnings:
            with st.expander(f"{len(result.warnings)} build warning(s)", expanded=True):
                for warning in result.warnings:
                    st.warning(warning)

        if result.pdf:
            st.caption(f"{result.pages} page(s) · "
                       f"{(st.session_state.result_keys.get(active) or 'portrait').title()} template")
            previews = _thumbs(result.pdf, 9)
            for start in range(0, len(previews), 3):
                columns = st.columns(3)
                for offset, image in enumerate(previews[start:start + 3]):
                    columns[offset].image(image, caption=f"Page {start + offset + 1}",
                                          use_container_width=True)
        else:
            st.info("Preview needs LibreOffice. The .docx download is ready.")

    if len(order) > 1:
        st.divider()
        built = [digest for digest in order if digest in st.session_state.results]
        st.subheader(f"Batch downloads ({len(built)}/{len(order)} built)")
        if built:
            st.download_button(
                f"⬇ Download all {len(built)} document(s) (.zip)",
                _zip_downloads(built), "SPT-documents.zip", "application/zip",
                use_container_width=True)
        for digest in order:
            res = st.session_state.results.get(digest)
            name = st.session_state.names.get(digest, digest[:8])
            row = st.columns([4, 1.2, 1.2])
            marker = "✅" if res else ("⚠️" if digest in st.session_state.build_errors else "·")
            row[0].write(f"{marker} {name}")
            if res is None:
                row[1].caption("not built")
                continue
            base = _base_name(digest)
            row[1].download_button("docx", res.docx, f"{base}-SPT.docx",
                                   "application/vnd.openxmlformats-officedocument."
                                   "wordprocessingml.document",
                                   key=f"dl_docx_{digest}", use_container_width=True)
            if res.pdf:
                row[2].download_button("pdf", res.pdf, f"{base}-SPT.pdf",
                                       "application/pdf", key=f"dl_pdf_{digest}",
                                       use_container_width=True)
            else:
                row[2].caption("no PDF")
