"""Dispatch any supported manuscript file to the right reader."""
from __future__ import annotations

import glob
import io
import os
import shutil
import subprocess
import tempfile
import zipfile

from . import model as M
from .read_docx import read_docx
from .read_pdf import read_pdf

SUPPORTED = ["docx", "doc", "odt", "rtf", "pdf", "tex", "zip", "md", "txt"]


def _run(cmd, cwd=None, timeout=180):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(f"{cmd[0]} failed: {r.stderr[-800:]}")
    return r


def _office_to_docx(data: bytes, ext: str) -> bytes:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(f".{ext} needs LibreOffice installed on the server; save the file as .docx and upload again.")
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, f"in.{ext}")
        open(src, "wb").write(data)
        _run([soffice, "--headless", "--convert-to", "docx", "--outdir", td, src])
        return open(os.path.join(td, "in.docx"), "rb").read()


def _pandoc_to_docx(workdir: str, main: str) -> bytes:
    if not shutil.which("pandoc"):
        raise RuntimeError("LaTeX/Markdown input needs pandoc installed on the server.")
    out = os.path.join(workdir, "__out.docx")
    cmd = ["pandoc", main, "-o", out, "--resource-path", workdir]
    bibs = glob.glob(os.path.join(workdir, "**", "*.bib"), recursive=True)
    if bibs:
        cmd += ["--citeproc"] + [f"--bibliography={b}" for b in bibs]
    _run(cmd, cwd=os.path.dirname(main) or workdir)
    return open(out, "rb").read()


def _find_main_tex(root: str) -> str:
    cands = glob.glob(os.path.join(root, "**", "*.tex"), recursive=True)
    for c in cands:
        try:
            if "\\documentclass" in open(c, encoding="utf-8", errors="ignore").read():
                return c
        except OSError:
            pass
    if cands:
        return cands[0]
    raise RuntimeError("No .tex file found in the upload.")


def read_any(filename: str, data: bytes) -> M.Document:
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext == "docx":
        doc = read_docx(io.BytesIO(data))
    elif ext in ("doc", "odt", "rtf"):
        doc = read_docx(io.BytesIO(_office_to_docx(data, ext)))
    elif ext == "pdf":
        doc = read_pdf(data)
    elif ext in ("tex", "md", "txt", "zip"):
        with tempfile.TemporaryDirectory() as td:
            if ext == "zip":
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    z.extractall(td)
                main = _find_main_tex(td)
            else:
                main = os.path.join(td, "main." + ("markdown" if ext in ("md", "txt") else "tex"))
                open(main, "wb").write(data)
            protected, prem = [], []
            if main.endswith(".tex"):
                from .texblocks import expand_inputs, protect
                src = open(main, encoding="utf-8", errors="ignore").read()
                src = expand_inputs(src, os.path.dirname(main))
                mod, protected, prem = protect(src)
                open(main, "w", encoding="utf-8").write(mod)
            doc = read_docx(io.BytesIO(_pandoc_to_docx(td, main)))
            if protected:
                from .texblocks import restore
                restore(doc, protected, prem)
        doc.source_type = ext
        if ext in ("tex", "zip"):
            doc.warnings.append("LaTeX input converted via pandoc — custom macros/environments may need a check.")
    else:
        raise RuntimeError(f"Unsupported file type .{ext}. Supported: {', '.join(SUPPORTED)}")
    doc.source_type = doc.source_type or ext
    return doc
