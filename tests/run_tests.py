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
sys.exit(1 if fails else 0)
