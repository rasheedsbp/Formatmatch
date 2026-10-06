"""Regression run: every input type × every preset, plus template mode. Usage: python tests/run_tests.py"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from formatmatch.engine import FormatSource, build_spec, format_manuscript, learn_template  # noqa: E402
from formatmatch.render_latex import try_compile  # noqa: E402
from formatmatch.spec import PRESETS  # noqa: E402

import make_fixtures  # noqa: E402

make_fixtures.make_apa_draft(os.path.join(HERE, "draft_apa.docx"))
make_fixtures.make_template(os.path.join(HERE, "template_journal.docx"))
draft = open(os.path.join(HERE, "draft_apa.docx"), "rb").read()
fails = 0
for key in PRESETS:
    res = format_manuscript("draft_apa.docx", draft, build_spec(FormatSource(preset=key)))
    ok, _, log = try_compile(res.latex_zip)
    refs_ok = len(res.report.refs) == 4
    status = ok and refs_ok and res.docx
    fails += not status
    print(f"{'PASS' if status else 'FAIL'}  {key:18s} latex={ok} refs={len(res.report.refs)} warnings={len(res.report.warnings)}")

ov, prof, _ = learn_template("template_journal.docx", open(os.path.join(HERE, "template_journal.docx"), "rb").read())
res = format_manuscript("draft_apa.docx", draft, build_spec(FormatSource(template_overrides=ov, profile=prof)), prof)
print(f"{'PASS' if res.docx else 'FAIL'}  template mode ({len(res.report.changes)} changes)")
tex = os.path.join(HERE, "texproj.zip")
if os.path.exists(tex):
    res = format_manuscript("texproj.zip", open(tex, "rb").read(), build_spec(FormatSource(preset="ieee_conference")))
    print(f"{'PASS' if len(res.report.refs) == 2 else 'FAIL'}  LaTeX input")

# ---- equations & algorithms (verbatim preservation)
import zipfile as _zf, io as _io  # noqa: E402
make_fixtures.make_math_algo(os.path.join(HERE, "math_algo.docx"))
src = open(os.path.join(HERE, "math_algo.docx"), "rb").read()
for key in ("ieee_journal", "springer_lncs", "mdpi"):
    res = format_manuscript("math_algo.docx", src, build_spec(FormatSource(preset=key)))
    xml = _zf.ZipFile(_io.BytesIO(res.docx)).read("word/document.xml").decode()
    ok = (xml.count("<m:oMath") >= 4 and xml.count("Equation.DSMT4") == 2 and "Algorithm 1:" in xml
          and "Algorithm 2: Genetic" in xml and "for each particle" in xml)
    lok, _, _ = try_compile(res.latex_zip)
    fails += not (ok and lok)
    print(f"{'PASS' if ok and lok else 'FAIL'}  maths/algorithms docx → {key} (latex={lok})")
tex_alg = os.path.join(HERE, "texalg.zip")
if os.path.exists(tex_alg):
    res = format_manuscript("texalg.zip", open(tex_alg, "rb").read(), build_spec(FormatSource(preset="ieee_journal")))
    main = _zf.ZipFile(_io.BytesIO(res.latex_zip)).read("main.tex").decode()
    ok = "\\begin{algorithmic}[1]" in main and "\\label{eq:energy}" in main and "\\begin{align}" in main
    fails += not ok
    print(f"{'PASS' if ok else 'FAIL'}  LaTeX algorithm/equation environments preserved verbatim")
sys.exit(1 if fails else 0)
