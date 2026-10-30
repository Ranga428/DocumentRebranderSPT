import base64
import io
import fitz
import streamlit as st
import converter as c

st.set_page_config(page_title="VAPT Report Converter", page_icon="SapientLogoNoText.png", layout="wide")
with open("SapientLogoNoText.png", "rb") as logo_file:
    logo_data = base64.b64encode(logo_file.read()).decode("ascii")
st.markdown(
    f"""
    <div style="display: flex; align-items: center; gap: 1rem;">
        <img src="data:image/png;base64,{logo_data}"
             alt="Sapient logo" style="display: block; width: 48px; height: 48px;">
        <h1 style="margin: 0; font-size: 2.75rem; line-height: 1.1;">
            VAPT Report Converter
        </h1>
    </div>
    """,
    unsafe_allow_html=True,
)
st.write("Upload the original report (PDF) and your template (the converted .docx). "
         "The app extracts the report data, applies your rewrite rules, and fills the template.")

c1, c2 = st.columns(2)
pdf_f = c1.file_uploader("1. Original report (PDF)", type="pdf")
tpl_f = c2.file_uploader("2. Template / converted sample (.docx)", type="docx")

st.sidebar.header("Settings")
documented_by = st.sidebar.text_input("Documented By", "Jenel Esteron")
asset_label = st.sidebar.text_input("Asset label (blank = auto)", "")

if pdf_f and tpl_f:
    pdf_b, tpl_b = pdf_f.getvalue(), tpl_f.getvalue()
    cfg = {"documented_by": documented_by, "asset_label": asset_label or None}
    rep = c.parse_original(pdf_b)
    plan = c.make_plan(rep, c.learn_library(tpl_b), cfg)

    st.subheader("Review")
    for w in plan["warnings"]:
        st.warning(w)
    for t in plan["images_missing"]:
        st.warning(f"No screenshot found for: {t}")
    a, b, d = st.columns(3)
    a.metric("Findings", len(plan["findings"]))
    b.metric("Assessment ID", plan["assessment_id"])
    d.metric("Dates", f"{plan['start']} – {plan['end']}")
    st.dataframe([{"Finding": f, "Field": k, "Source": s} for f, k, s in plan["review"]],
                 use_container_width=True)

    for i, f in enumerate(plan["findings"]):
        with st.expander(f"Edit: {f['title']}"):
            for key in ("description", "corrective", "reference", "details", "hosts", "steps"):
                txt = st.text_area(key.title() + " (one paragraph / item per line)",
                                   "\n".join(f[key]), key=f"{i}{key}", height=120)
                f[key] = [x for x in txt.split("\n") if x.strip()]

    if st.button("Build report", type="primary"):
        with st.spinner("Building..."):
            data, warns = c.build_docx(plan, tpl_b)
            pdf = c.render_pdf(data)
        for w in warns[len(plan["warnings"]):]:
            st.warning(w)
        st.download_button("📥 Download .docx", data, f"report_{plan['assessment_id']}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if pdf:
            st.download_button("📥 Download PDF preview", pdf, f"report_{plan['assessment_id']}.pdf", "application/pdf")
            doc = fitz.open(stream=pdf, filetype="pdf")
            cols = st.columns(4)
            for i, p in enumerate(doc):
                cols[i % 4].image(p.get_pixmap(dpi=60).tobytes("png"), caption=f"Page {i + 1}")