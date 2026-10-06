"""High-level API: build the target spec from all sources and format a manuscript."""
from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Optional

from .pipeline import Report, prepare
from .read_any import read_any
from .render_docx import render_docx
from .render_latex import render_latex
from .spec import DEFAULT_SPEC, deep_merge, get_preset, lookup_journal, parse_guidelines, parse_guidelines_llm
from .template import analyze_docx_template, analyze_pdf_sample


def dotx_to_docx(data: bytes) -> bytes:
    """Word templates (.dotx/.dotm) differ from .docx only in the main part's content type."""
    src = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                content = re.sub(rb"application/vnd\.ms-word\.template\.macroEnabledTemplate\.main\+xml|"
                                 rb"application/vnd\.openxmlformats-officedocument\.wordprocessingml\.template\.main\+xml",
                                 b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                                 content)
            out.writestr(item, content)
    return buf.getvalue()


def guidelines_text_from_file(name: str, data: bytes) -> str:
    ext = name.rsplit(".", 1)[-1].lower()
    if ext == "pdf":
        import pymupdf
        return "\n".join(p.get_text() for p in pymupdf.open(stream=data, filetype="pdf"))
    if ext in ("docx", "dotx"):
        import docx
        d = docx.Document(io.BytesIO(dotx_to_docx(data) if ext == "dotx" else data))
        return "\n".join(p.text for p in d.paragraphs)
    if ext in ("html", "htm"):
        return html_to_text(data.decode("utf-8", "ignore"))
    return data.decode("utf-8", "ignore")


def html_to_text(html: str) -> str:
    html = re.sub(r"(?is)<(script|style|nav|footer|header)[^>]*>.*?</\1>", " ", html)
    html = re.sub(r"(?i)<br\s*/?>|</p>|</li>|</h\d>|</div>", "\n", html)
    text = re.sub(r"<[^>]+>", " ", html)
    import html as h
    text = h.unescape(text)
    return re.sub(r"[ \t]+", " ", re.sub(r"\n\s*\n+", "\n\n", text)).strip()


def fetch_guidelines_url(url: str, timeout=20) -> str:
    import requests
    r = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0 (FormatMatch guideline reader)"})
    r.raise_for_status()
    ctype = r.headers.get("content-type", "")
    if "pdf" in ctype or url.lower().endswith(".pdf"):
        return guidelines_text_from_file("g.pdf", r.content)
    return html_to_text(r.text)


@dataclass
class FormatSource:
    preset: Optional[str] = None                 # preset key chosen/derived from journal name
    journal: Optional[dict] = None               # lookup_journal() result
    guidelines: Optional[dict] = None            # overrides from guidelines
    template_overrides: Optional[dict] = None    # overrides from template/sample
    profile: object = None                       # TemplateProfile (docx templates only)
    manual: dict = field(default_factory=dict)   # user edits in the UI
    found: list = field(default_factory=list)    # human-readable detection notes


def build_spec(src: FormatSource) -> dict:
    spec = get_preset(src.preset) if src.preset else DEFAULT_SPEC
    if src.guidelines:
        spec = deep_merge(spec, {k: v for k, v in src.guidelines.items() if not k.startswith("_")})
    if src.template_overrides:
        spec = deep_merge(spec, {k: v for k, v in src.template_overrides.items() if not k.startswith("_")})
    if src.manual:
        spec = deep_merge(spec, src.manual)
    if src.journal and src.journal.get("title"):
        spec["name"] = src.journal["title"] + (f" ({spec.get('name')})" if spec.get("name") else "")
    return spec


def learn_template(name: str, data: bytes):
    """Returns (overrides, profile_or_None, found_notes)."""
    ext = name.rsplit(".", 1)[-1].lower()
    if ext in ("dotx", "dotm"):
        data = dotx_to_docx(data)
        ext = "docx"
    if ext == "docx":
        ov, prof = analyze_docx_template(data)
        return ov, prof, list(prof.found)
    if ext == "pdf":
        ov = analyze_pdf_sample(data)
        return ov, None, ov.pop("_found", [])
    if ext in ("doc", "odt", "rtf"):
        from .read_any import _office_to_docx
        d2 = _office_to_docx(data, ext)
        ov, prof = analyze_docx_template(d2)
        return ov, prof, list(prof.found)
    raise RuntimeError("Template/sample must be .docx, .dotx, .doc or a published-paper .pdf")


@dataclass
class Result:
    docx: Optional[bytes]
    latex_zip: Optional[bytes]
    report: Report
    report_md: str
    spec: dict
    preview_pdf: Optional[bytes] = None


def format_manuscript(filename: str, data: bytes, spec: dict, profile=None, want_docx=True, want_latex=True,
                      crossref=False, progress=None, preview=False) -> Result:
    def prog(f, msg):
        if progress:
            progress(f, msg)

    prog(0.05, "Reading manuscript")
    doc = read_any(filename, data)
    prog(0.2, "Restructuring: headings, captions, citations, references")
    prepared, rep = prepare(doc, spec, crossref=crossref,
                            progress=(lambda f, m: prog(0.2 + 0.4 * f, m)) if crossref else None)
    docx_bytes = latex_bytes = None
    if want_docx:
        prog(0.65, "Writing Word document")
        docx_bytes, w = render_docx(prepared, spec, profile)
        for x in w:
            rep.warn(x)
    if want_latex:
        prog(0.8, "Writing LaTeX project")
        lw = []
        latex_bytes = render_latex(prepared, spec, rep.refs, lw)
        for x in lw:
            rep.warn(x)
    pdf = None
    if preview and docx_bytes:
        prog(0.9, "Rendering preview")
        pdf = docx_to_pdf(docx_bytes)
    prog(1.0, "Done")
    return Result(docx_bytes, latex_bytes, rep, report_markdown(rep, spec, filename), spec, pdf)


def docx_to_pdf(data: bytes) -> Optional[bytes]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "out.docx")
        open(src, "wb").write(data)
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, src],
                           capture_output=True, timeout=180)
        except Exception:
            return None
        pdf = os.path.join(td, "out.pdf")
        return open(pdf, "rb").read() if os.path.exists(pdf) else None


def report_markdown(rep: Report, spec: dict, filename: str) -> str:
    L = [f"# Formatting report — {filename}", "", f"**Target format:** {spec.get('name', 'Custom')}", ""]
    L += ["## Compliance checks", "", "| Check | Status | Detail |", "|---|---|---|"]
    for name, ok, detail in rep.checks:
        L.append(f"| {name} | {'✅ Pass' if ok else '⚠️ Fix'} | {detail} |")
    L += ["", "## Changes made", ""] + [f"- {c}" for c in rep.changes]
    if rep.warnings:
        L += ["", "## Needs your attention", ""] + [f"- {w}" for w in rep.warnings]
    S = spec
    L += ["", "## Format applied", "",
          f"- Page: {S['page']['size']}, {S['page']['columns']} column(s), margins (T/B/L/R) {S['page']['margins_cm']} cm",
          f"- Body: {S['font']['family']} {S['font']['size']} pt, line spacing {S['font']['line_spacing']}",
          f"- Headings: {S['headings']['scheme']} numbering",
          f"- References: {S['references']['style']}; citations: {S['citations']['mode']}",
          "", "_Generated by FormatMatch — CS Software Solutions (cssoftwaresolutions.in)_"]
    return "\n".join(L)
