"""FormatMatch — arrange any manuscript into a target journal / template format.

Run:  streamlit run app.py
"""
import io
import os

import streamlit as st

from formatmatch.engine import (FormatSource, build_spec, fetch_guidelines_url, format_manuscript,
                                guidelines_text_from_file, learn_template)
from formatmatch.read_any import SUPPORTED
from formatmatch.spec import PRESETS, get_preset, lookup_journal, parse_guidelines, parse_guidelines_llm

st.set_page_config(page_title="FormatMatch · Manuscript Formatter", page_icon="📄", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.6rem; max-width: 1180px;}
h1 {font-weight: 750; letter-spacing: -0.02em;}
.fm-sub {color: #5b6472; margin-top: -0.6rem; margin-bottom: 1.2rem;}
.fm-step {font-size: 0.78rem; font-weight: 700; letter-spacing: .08em; color: #1f6feb; text-transform: uppercase;}
.fm-found {background: #f3f6fb; border-left: 3px solid #1f6feb; padding: .55rem .8rem; border-radius: 4px;
           font-size: .9rem; margin: .3rem 0;}
@media (prefers-color-scheme: dark) { .fm-found {background: #161b22;} .fm-sub {color: #9aa4b2;} }
.fm-foot {color: #8a93a0; font-size: .82rem; margin-top: 2rem; border-top: 1px solid #e3e7ee; padding-top: .8rem;}
</style>
""", unsafe_allow_html=True)

st.title("FormatMatch")
st.markdown('<div class="fm-sub">Upload a manuscript, give the target format (template, author guidelines or journal '
            'name) and get it back arranged exactly to that format — Word and LaTeX, with a compliance report.</div>',
            unsafe_allow_html=True)

PRESET_KEYS = list(PRESETS.keys())
PRESET_LABEL = {k: PRESETS[k]["name"] for k in PRESET_KEYS}
REF_STYLES = ["ieee", "apa", "harvard", "emerald-harvard", "vancouver", "ama", "chicago", "springer-basic",
              "springer-vancouver", "lncs", "mdpi", "elsevier-numeric", "mla"]
CITE_MODES = ["numeric", "superscript", "author-year", "author-year-nocomma"]
FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Palatino Linotype", "Garamond", "Georgia",
         "Book Antiqua", "Helvetica", "Computer Modern"]

ss = st.session_state
ss.setdefault("found", {})

# ------------------------------------------------------------------ step 1
st.markdown('<div class="fm-step">Step 1 · Manuscript</div>', unsafe_allow_html=True)
ms = st.file_uploader("Manuscript to format", type=SUPPORTED, key="ms",
                      help="Word (.docx/.doc/.odt/.rtf), PDF, LaTeX (.tex, or a .zip with .tex + .bib + figures), Markdown")

# ------------------------------------------------------------------ step 2
st.markdown('<div class="fm-step">Step 2 · Target format</div>', unsafe_allow_html=True)
st.caption("Use any combination. Priority when they disagree: your edits › template/sample › guidelines › journal preset.")
tab_t, tab_g, tab_j = st.tabs(["📄 Template / sample paper", "📋 Author guidelines", "🔎 Journal name"])

src = FormatSource()

with tab_t:
    tf = st.file_uploader("Journal template (.docx/.dotx) or a published sample paper (.pdf)",
                          type=["docx", "dotx", "dotm", "doc", "pdf"], key="tpl")
    if tf is not None:
        try:
            if ss.get("tpl_name") != tf.name + str(tf.size):
                ov, prof, found = learn_template(tf.name, tf.getvalue())
                ss.tpl_cache = (ov, prof, found)
                ss.tpl_name = tf.name + str(tf.size)
            ov, prof, found = ss.tpl_cache
            src.template_overrides, src.profile = ov, prof
            st.markdown("".join(f'<div class="fm-found">✓ {f}</div>' for f in found), unsafe_allow_html=True)
            if prof is not None:
                st.success("Word template: styles, fonts, spacing, columns, headers/footers are cloned exactly from it.")
            else:
                st.info("Sample PDF: layout and style rules were measured from the paper.")
        except Exception as e:
            st.error(f"Could not read the template: {e}")

with tab_g:
    gl_text = st.text_area("Paste the author guidelines (or the relevant parts)", height=160, key="gl_text")
    c1, c2 = st.columns(2)
    gl_file = c1.file_uploader("…or upload them (PDF/DOCX/TXT/HTML)", type=["pdf", "docx", "txt", "html", "htm", "md"],
                               key="gl_file")
    gl_url = c2.text_input("…or a guidelines URL", key="gl_url", placeholder="https://…/author-guidelines")
    use_llm = c2.checkbox("Use AI to read the guidelines (Groq)", value=False,
                          help="Rule-based reading works offline. AI reading handles unusual wording; needs a free Groq API key.")
    groq_key = ""
    if use_llm:
        try:
            groq_key = st.secrets.get("GROQ_API_KEY", "")
        except Exception:
            groq_key = ""
        if not groq_key:
            groq_key = c2.text_input("Groq API key", type="password", key="groq")
    text = gl_text or ""
    if gl_file is not None:
        text += "\n" + guidelines_text_from_file(gl_file.name, gl_file.getvalue())
    if gl_url:
        if ss.get("gl_url_done") != gl_url:
            try:
                ss.gl_url_text = fetch_guidelines_url(gl_url)
                ss.gl_url_done = gl_url
            except Exception as e:
                st.warning(f"Could not fetch that page ({e}). Paste the text instead.")
                ss.gl_url_text = ""
        text += "\n" + ss.get("gl_url_text", "")
    if text.strip():
        try:
            g = parse_guidelines_llm(text, groq_key) if (use_llm and groq_key) else parse_guidelines(text)
        except Exception as e:
            st.warning(f"AI reading failed ({e}); using rule-based reading.")
            g = parse_guidelines(text)
        src.guidelines = g
        found = g.get("_found", [])
        if found:
            st.markdown("".join(f'<div class="fm-found">✓ {f}</div>' for f in found), unsafe_allow_html=True)
        else:
            st.info("No explicit formatting rules recognised — the journal preset / template will be used.")

with tab_j:
    jn = st.text_input("Target journal or conference", placeholder="e.g. IEEE Access, Expert Systems with Applications, "
                                                                   "International Journal of Organizational Analysis")
    if jn:
        if ss.get("jn_done") != jn:
            with st.spinner("Looking up the journal…"):
                ss.jn_res = lookup_journal(jn)
                ss.jn_done = jn
        res = ss.jn_res
        src.journal = res
        if res.get("publisher"):
            st.markdown(f'<div class="fm-found">✓ {res["title"]} — {res["publisher"]}'
                        f'{" · ISSN " + res["issn"] if res.get("issn") else ""}</div>', unsafe_allow_html=True)
        if not res.get("preset"):
            st.warning("Publisher not recognised automatically — choose the closest house style below.")
    default_idx = PRESET_KEYS.index(ss.jn_res["preset"]) if jn and ss.get("jn_res", {}).get("preset") else None
    choice = st.selectbox("House style", PRESET_KEYS, index=default_idx, format_func=lambda k: PRESET_LABEL[k],
                          placeholder="Choose a publisher style (optional if you gave a template)")
    src.preset = choice

# ------------------------------------------------------------------ step 3
st.markdown('<div class="fm-step">Step 3 · Review & fine-tune</div>', unsafe_allow_html=True)
spec0 = build_spec(src)
with st.expander("Detected format — edit anything before formatting", expanded=False):
    man = {}
    a, b, c, d = st.columns(4)
    pg, ft = spec0["page"], spec0["font"]
    v = a.selectbox("Paper", ["A4", "Letter"], index=0 if pg["size"] == "A4" else 1)
    man.setdefault("page", {})["size"] = v
    man["page"]["columns"] = b.selectbox("Columns", [1, 2], index=int(pg["columns"]) - 1)
    m = pg["margins_cm"]
    mv = c.number_input("Margins (cm, all sides)", 0.5, 6.0, float(round(sum(m) / 4, 2)), 0.05)
    if abs(mv - sum(m) / 4) > 0.01:
        man["page"]["margins_cm"] = [mv] * 4
    man["page"]["line_numbers"] = d.checkbox("Line numbers", value=bool(pg.get("line_numbers")))
    a, b, c, d = st.columns(4)
    fam = ft["family"]
    man.setdefault("font", {})["family"] = a.selectbox("Body font", FONTS + ([fam] if fam not in FONTS else []),
                                                       index=(FONTS + [fam]).index(fam))
    man["font"]["size"] = b.number_input("Font size (pt)", 7.0, 16.0, float(ft["size"]), 0.5)
    man["font"]["line_spacing"] = c.selectbox("Line spacing", [1.0, 1.15, 1.5, 2.0],
                                              index=[1.0, 1.15, 1.5, 2.0].index(ft["line_spacing"]) if ft["line_spacing"] in [1.0, 1.15, 1.5, 2.0] else 0)
    man["font"]["align"] = d.selectbox("Alignment", ["justify", "left"], index=0 if ft.get("align") == "justify" else 1)
    a, b, c, d = st.columns(4)
    hs = spec0["headings"]["scheme"]
    man["headings"] = {"scheme": a.selectbox("Heading numbering", ["decimal", "ieee", "none"],
                                             index=["decimal", "ieee", "none"].index(hs),
                                             format_func=lambda x: {"decimal": "1. / 1.1", "ieee": "I. / A. / 1)", "none": "Unnumbered"}[x])}
    rs = spec0["references"]["style"]
    man["references"] = {"style": b.selectbox("Reference style", REF_STYLES, index=REF_STYLES.index(rs) if rs in REF_STYLES else 0)}
    cm = spec0["citations"]["mode"]
    man["citations"] = {"mode": c.selectbox("In-text citations", CITE_MODES, index=CITE_MODES.index(cm) if cm in CITE_MODES else 0)}
    lim = spec0["limits"]
    aw = d.number_input("Abstract word limit (0 = none)", 0, 1000, int(lim.get("abstract_words") or 0), 10)
    man["limits"] = {"abstract_words": aw or None}
    a, b, c, d = st.columns(4)
    man["abstract"] = {"layout": a.selectbox("Abstract layout", ["block", "inline"],
                                             index=0 if spec0["abstract"]["layout"] == "block" else 1),
                       "structured": b.checkbox("Structured abstract", value=bool(spec0["abstract"].get("structured")))}
    man["keywords"] = {"label": c.text_input("Keywords label", spec0["keywords"]["label"]),
                       "separator": d.selectbox("Keyword separator", ["; ", ", ", " · "],
                                                index=["; ", ", ", " · "].index(spec0["keywords"]["separator"]) if spec0["keywords"]["separator"] in ["; ", ", ", " · "] else 0,
                                                format_func=lambda s: repr(s.strip()))}
    a, b, c, d = st.columns(4)
    man["captions"] = {"figure_label": a.text_input("Figure label", spec0["captions"]["figure_label"]),
                       "table_label": b.text_input("Table label", spec0["captions"]["table_label"]),
                       "table_numbering": c.selectbox("Table numbers", ["arabic", "roman"],
                                                      index=0 if spec0["captions"].get("table_numbering") != "roman" else 1)}
    man["references"]["heading"] = d.text_input("Reference heading", spec0["references"].get("heading", "References"))
    if src.profile is not None:
        st.caption("Word output clones the template's own styles; these settings mainly steer numbering, labels, "
                   "references and the LaTeX output.")
    src.manual = man
spec = build_spec(src)
with st.expander("Advanced: full format specification (JSON)"):
    st.json(spec, expanded=False)

# ------------------------------------------------------------------ step 4
st.markdown('<div class="fm-step">Step 4 · Format</div>', unsafe_allow_html=True)
a, b, c, d = st.columns(4)
want_docx = a.checkbox("Word (.docx)", value=True)
want_latex = b.checkbox("LaTeX project (.zip)", value=True)
crossref = c.checkbox("Verify references online (Crossref)", value=False,
                      help="Completes missing volume/issue/pages/DOI and fixes author names from Crossref metadata.")
preview = d.checkbox("Show preview", value=True)

ready = ms is not None and (src.preset or src.template_overrides or src.guidelines)
if ms is not None and not ready:
    st.info("Give at least one target format in Step 2 (template, guidelines or journal style).")
go = st.button("Format manuscript", type="primary", disabled=not ready, width="stretch")

if go:
    bar = st.progress(0.0, text="Starting…")
    try:
        res = format_manuscript(ms.name, ms.getvalue(), spec, src.profile, want_docx, want_latex, crossref,
                                progress=lambda f, m: bar.progress(min(1.0, f), text=m), preview=preview)
        ss.result = res
        ss.result_name = os.path.splitext(ms.name)[0]
    except Exception as e:
        bar.empty()
        st.error(f"Formatting failed: {e}")
        st.stop()

res = ss.get("result")
if res is not None:
    base = ss.get("result_name", "manuscript")
    st.divider()
    checks = res.report.checks
    n_ok = sum(1 for _, ok, _ in checks if ok)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Checks passed", f"{n_ok}/{len(checks)}")
    k2.metric("Changes applied", len(res.report.changes))
    k3.metric("Need attention", len(res.report.warnings))
    k4.metric("References", len(res.report.refs))
    a, b, c = st.columns(3)
    if res.docx:
        a.download_button("⬇ Word document", res.docx, file_name=f"{base}_formatted.docx", width="stretch",
                          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if res.latex_zip:
        b.download_button("⬇ LaTeX project", res.latex_zip, file_name=f"{base}_latex.zip", width="stretch",
                          mime="application/zip")
    c.download_button("⬇ Formatting report", res.report_md.encode(), file_name=f"{base}_format_report.md",
                      width="stretch", mime="text/markdown")
    left, right = st.columns([1.05, 1])
    with left:
        st.subheader("Compliance")
        for name, ok, detail in checks:
            st.markdown(f"{'✅' if ok else '⚠️'} **{name}** — {detail}")
        if res.report.warnings:
            st.subheader("Needs your attention")
            for w in res.report.warnings:
                st.markdown(f"- {w}")
        st.subheader("Changes made")
        for ch in res.report.changes:
            st.markdown(f"- {ch}")
    with right:
        if res.preview_pdf:
            st.subheader("Preview")
            import pymupdf
            pdf = pymupdf.open(stream=res.preview_pdf, filetype="pdf")
            pages = len(pdf)
            pno = st.number_input(f"Page (of {pages})", 1, pages, 1) if pages > 1 else 1
            pix = pdf[pno - 1].get_pixmap(dpi=90)
            st.image(pix.tobytes("png"), width="stretch")
            st.caption("Preview rendered with LibreOffice — Word may differ slightly (equations render in Word).")
        elif preview:
            st.caption("Preview needs LibreOffice on the server; download the Word file to view it.")

st.markdown('<div class="fm-foot">FormatMatch · CS Software Solutions · '
            '<a href="https://cssoftwaresolutions.in">cssoftwaresolutions.in</a> · info@cssoftwaresolutions.tech · '
            'Manuscript editing, statistics and publication support</div>', unsafe_allow_html=True)
