"""Compliance checks against the format's limits and requirements."""
from __future__ import annotations

import re

from . import model as M
from .classify import int_to_roman


def _words(s: str) -> int:
    return len(re.findall(r"\b[\w'’-]+\b", s))


def run_checks(doc: M.Document, spec: dict, rep):
    L = spec.get("limits", {})
    add = lambda name, ok, detail: rep.checks.append((name, ok, detail))
    title = doc.first(M.TITLE)
    if title is None:
        add("Title present", False, "No title detected")
    elif L.get("title_words"):
        n = _words(title.text)
        add("Title length", n <= L["title_words"], f"{n} words (limit {L['title_words']})")
    abs_blocks = doc.by_role(M.ABSTRACT)
    if not abs_blocks:
        add("Abstract present", False, "No abstract detected")
    else:
        n = sum(_words(b.text) for b in abs_blocks)
        if L.get("abstract_words"):
            add("Abstract length", n <= L["abstract_words"], f"{n} words (limit {L['abstract_words']})")
        else:
            add("Abstract present", True, f"{n} words")
        if spec["abstract"].get("structured"):
            have = {b.label.lower() for b in abs_blocks if b.label}
            add("Structured abstract", bool(have), ", ".join(sorted(have)) or "not structured")
    kw = doc.first(M.KEYWORDS)
    if kw is None:
        add("Keywords present", False, "No keywords line detected")
    else:
        n = int(kw.label or 0)
        lo, hi = L.get("keywords_min"), L.get("keywords_max")
        ok = (lo is None or n >= lo) and (hi is None or n <= hi)
        rng = f"{lo or 0}–{hi}" if hi else (f"≥{lo}" if lo else "no limit")
        add("Keyword count", ok, f"{n} keywords (required {rng})")
    if L.get("total_words"):
        n = sum(_words(b.text) for b in doc.blocks if b.role in (M.PARA, M.LIST, M.HEADING, M.ABSTRACT))
        add("Manuscript length", n <= L["total_words"], f"{n} words excl. references (limit {L['total_words']})")
    heads = [b.text.lower() for b in doc.blocks if b.role == M.HEADING]
    for sec in spec.get("required_sections", []):
        ok = any(sec.lower()[:6] in h for h in heads)
        add(f"Section “{sec}”", ok, "present" if ok else "missing — required by the guidelines")
    # floats
    figs = [b for b in doc.blocks if b.role == M.FIGURE]
    caps = [b for b in doc.blocks if b.role == M.CAPTION]
    tabs = [b for b in doc.blocks if b.role == M.TABLE]
    fcap = [c for c in caps if c.cap_kind == "figure"]
    tcap = [c for c in caps if c.cap_kind == "table"]
    if figs:
        add("Figure captions", len(fcap) >= 1, f"{len(figs)} image(s), {len(fcap)} figure caption(s)")
    if tabs:
        add("Table captions", len(tcap) >= len(tabs), f"{len(tabs)} table(s), {len(tcap)} table caption(s)")
    body = " ".join(b.text for b in doc.blocks if b.role in (M.PARA, M.LIST))
    roman = spec["captions"].get("table_numbering") == "roman"
    uncited = []
    for c in caps:
        num = int_to_roman(c.cap_num) if (c.cap_kind == "table" and roman) else str(c.cap_num)
        pat = r"\b(?:Fig(?:ure)?s?\.?|FIG\.?)\s*" if c.cap_kind == "figure" else r"\bTables?\s*"
        if not re.search(pat + r"(?:\w+\s*(?:[-–,]|and)\s*)?" + re.escape(num) + r"\b", body, re.I):
            uncited.append(f"{'Fig.' if c.cap_kind == 'figure' else 'Table'} {num}")
    if caps:
        add("Figures/tables cited in text", not uncited, "all cited" if not uncited else "not cited: " + ", ".join(uncited))
    refs = doc.by_role(M.REFERENCE)
    add("Reference list", bool(refs), f"{len(refs)} references")
