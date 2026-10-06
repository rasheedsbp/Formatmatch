"""Normalise a parsed manuscript against a spec: numbering, captions, citations, references, keywords."""
from __future__ import annotations

import copy
import re
from typing import List, Optional, Tuple

from . import model as M
from .classify import UNNUMBERED, int_to_roman, known_section
from .model import Run
from .refs import (NUMERIC_STYLES, CitationEngine, detect_citation_mode, enrich_crossref, format_reference,
                   parse_reference, sort_author_year, _apply, _segments)
from .runs_util import apply_case, merge_adjacent


class Report:
    def __init__(self):
        self.changes: List[str] = []
        self.warnings: List[str] = []
        self.checks: List[Tuple[str, bool, str]] = []   # (check, ok, detail)
        self.found: List[str] = []
        self.refs = []          # parsed references in final order

    def change(self, s):
        self.changes.append(s)

    def warn(self, s):
        if s not in self.warnings:
            self.warnings.append(s)


# ------------------------------------------------------------------ captions & floats
def _pair_floats(blocks: List[M.Block], spec, rep: Report) -> List[M.Block]:
    """Attach each caption to its adjacent figure/table and order them per spec."""
    fig_pos = spec["captions"].get("figure_position", "below")
    tab_pos = spec["captions"].get("table_position", "above")
    out: List[M.Block] = []
    i = 0
    used = set()
    n = len(blocks)
    while i < n:
        b = blocks[i]
        if i in used:
            i += 1
            continue
        if b.role == M.CAPTION:
            want = M.FIGURE if b.cap_kind == "figure" else M.TABLE
            # look at neighbours (skip nothing): next, then previous already-emitted
            j = None
            if i + 1 < n and blocks[i + 1].role == want and (i + 1) not in used:
                j = i + 1
                k = j
                group = [blocks[j]]
                # multi-panel figures: consecutive images
                while want == M.FIGURE and k + 1 < n and blocks[k + 1].role == M.FIGURE:
                    k += 1
                    group.append(blocks[k])
                for x in range(j, k + 1):
                    used.add(x)
                pos = fig_pos if want == M.FIGURE else tab_pos
                out.extend([b] + group if pos == "above" else group + [b])
                i += 1
                continue
            if out and out[-1].role == want:
                # float came first; caption follows it
                g = []
                while out and out[-1].role == want:
                    g.insert(0, out.pop())
                pos = fig_pos if want == M.FIGURE else tab_pos
                out.extend([b] + g if pos == "above" else g + [b])
                i += 1
                continue
            out.append(b)
            i += 1
            continue
        out.append(b)
        i += 1
    return out


def _number_captions(blocks: List[M.Block], spec, rep: Report):
    cap = spec["captions"]
    maps = {"figure": {}, "table": {}}
    cnt = {"figure": 0, "table": 0}
    for b in blocks:
        if b.role != M.CAPTION:
            continue
        cnt[b.cap_kind] += 1
        new = cnt[b.cap_kind]
        if b.cap_num is not None and b.cap_num != new:
            rep.change(f"{b.cap_kind.title()} {b.cap_num} renumbered to {new} (order of appearance)")
        if b.cap_num is not None:
            maps[b.cap_kind].setdefault(b.cap_num, new)
        b.cap_num = new
        if b.cap_kind == "figure":
            label = cap["figure_label"]
            num = str(new)
        else:
            label = cap["table_label"]
            if cap.get("table_label_case") == "upper":
                label = label.upper()
            num = int_to_roman(new) if cap.get("table_numbering") == "roman" else str(new)
        sep = cap.get("separator", ". ")
        if b.cap_kind == "table" and cap.get("table_label_newline"):
            sep = "\n"
        b.prefix = f"{label} {num}{sep}"
        # caption body formatting
        if b.cap_kind == "table" and cap.get("table_caption_case") == "upper":
            b.runs = [copy.copy(r) for r in b.runs]
            for r in b.runs:
                r.text = r.text.upper() if not r.is_math else r.text
        end_period = cap.get("table_end_period", cap.get("end_period", True)) if b.cap_kind == "table" \
            else cap.get("end_period", True)
        if b.runs and end_period:
            last = b.runs[-1]
            if not last.is_math and not last.text.rstrip().endswith((".", "?", "!")):
                b.runs = b.runs[:-1] + [Run(last.text.rstrip() + ".", last.bold, last.italic, last.sup, last.sub)]
        elif b.runs and not end_period:
            last = b.runs[-1]
            if not last.is_math and last.text.rstrip().endswith("."):
                b.runs = b.runs[:-1] + [Run(last.text.rstrip()[:-1], last.bold, last.italic, last.sup, last.sub)]
    return maps


XREF_RE = re.compile(r"\b(Figures?|Figs?\.?|FIGURES?|FIGS?\.?|Tables?|TABLES?)\s*(\d+|[IVXLC]+)\b(\s*(?:[-–]|and|&|,)\s*(\d+|[IVXLC]+)\b)?")


def _fix_xrefs(blocks: List[M.Block], spec, maps, rep: Report):
    xr = spec.get("body_xref", {})
    figw = xr.get("figure", spec["captions"]["figure_label"])
    tabw = xr.get("table", spec["captions"]["table_label"])
    roman = spec["captions"].get("table_numbering") == "roman"
    changed = 0

    def conv(kind, s):
        n = int(s) if s.isdigit() else None
        if n is None:
            from .classify import roman_to_int
            n = roman_to_int(s)
        n = maps[kind].get(n, n)
        return int_to_roman(n) if (kind == "table" and roman) else str(n)

    for b in blocks:
        if b.role not in (M.PARA, M.LIST, M.CAPTION):
            continue
        segs = _segments(b.runs)
        full = "".join(t for _, t, _ in segs)
        reps = []
        for m in XREF_RE.finditer(full):
            word = m.group(1)
            kind = "table" if word.lower().startswith("tab") else "figure"
            plural = word.lower().rstrip(".").endswith("s") or bool(m.group(3))
            if kind == "figure":
                lab = ("Figs." if figw.endswith(".") else figw + "s") if plural else figw
            else:
                lab = tabw + ("s" if plural else "")
            # sentence-initial "Fig." is usually spelled out ("Figure") in IEEE
            before = full[: m.start()].rstrip()
            if kind == "figure" and figw.endswith(".") and (not before or before.endswith((".", "?", "!"))):
                lab = "Figures" if plural else "Figure"
            new = f"{lab} {conv(kind, m.group(2))}"
            if m.group(3):
                joiner = m.group(3)[: m.group(3).index(m.group(4))]
                new += joiner + conv(kind, m.group(4))
            if new != m.group(0):
                reps.append((m.start(), m.end(), [Run(new)]))
        if reps:
            changed += len(reps)
            b.runs = _apply(segs, reps)
    if changed:
        rep.change(f"{changed} figure/table cross-references updated to “{figw} n” / “{tabw} n” style")


# ------------------------------------------------------------------ headings
def _number_headings(blocks: List[M.Block], spec, rep: Report):
    hs = spec["headings"]
    scheme = hs.get("scheme", "decimal")
    auto = set(hs.get("auto_levels", []))
    counters = [0, 0, 0, 0]
    in_appendix = False
    for b in blocks:
        if b.role != M.HEADING:
            continue
        lvl = max(1, min(b.level or 1, 4))
        b.level = lvl
        bare = b.text.strip()
        ks = known_section(bare)
        levels = hs.get("levels", [])
        case = levels[lvl - 1].get("case", "as-is") if lvl - 1 < len(levels) else "as-is"
        if case != "as-is":
            b.runs = [Run(apply_case(bare, case))]
        if ks in UNNUMBERED or bare.lower().startswith("appendix"):
            if lvl == 1:
                in_appendix = bare.lower().startswith("appendix")
            b.prefix = ""
            continue
        if scheme == "none" or lvl in auto:
            b.prefix = ""
            continue
        counters[lvl - 1] += 1
        for k in range(lvl, 4):
            counters[k] = 0
        if scheme == "ieee":
            p = [int_to_roman(counters[0]) + ".", chr(64 + max(1, counters[1])) + ".",
                 f"{counters[2]})", chr(96 + max(1, counters[3])) + ")"][lvl - 1]
        else:
            nums = ".".join(str(max(1, c)) for c in counters[:lvl])
            dot = hs.get("decimal_trailing_dot", True) if lvl == 1 else hs.get("subsection_trailing_dot", True)
            p = nums + ("." if dot else "")
        b.prefix = p + " "
    rep.change(f"Section headings renumbered ({scheme} scheme) and cased per format")


# ------------------------------------------------------------------ keywords / title / abstract
def _keywords(blocks, spec, rep: Report):
    kw = spec["keywords"]
    for b in blocks:
        if b.role != M.KEYWORDS:
            continue
        t = b.text.strip().rstrip(".")
        parts = [p.strip() for p in re.split(r"\s*[;·•]\s*|\s*,\s*(?![^()]*\))|\s*\|\s*|\s+—\s+", t) if p.strip()]
        if len(parts) <= 1 and "  " in t:
            parts = [p.strip() for p in re.split(r"\s{2,}", t) if p.strip()]
        case = kw.get("case", "as-is")
        fixed = []
        for i, p in enumerate(parts):
            if case == "lower-first":
                # IEEE: first term capitalised, rest lower-case except acronyms/proper nouns
                p = p[:1].upper() + p[1:] if i == 0 else (p if (p.isupper() or any(c.isupper() for c in p[1:])) else p[:1].lower() + p[1:])
            elif case in ("title", "sentence", "upper"):
                p = apply_case(p, case)
            fixed.append(p)
        b.runs = [Run(kw.get("separator", "; ").join(fixed) + ("." if kw.get("end_period") else ""))]
        b.label = str(len(fixed))
        rep.change(f"Keywords re-formatted ({len(fixed)} terms, separator “{kw.get('separator', '; ').strip()}”)")


def _title(blocks, spec):
    case = spec["title"].get("case", "as-is")
    if case == "as-is":
        return
    for b in blocks:
        if b.role == M.TITLE:
            if all(not r.is_math for r in b.runs) and not any(r.italic for r in b.runs):
                b.runs = [Run(apply_case(b.text.strip(), case))]
            break


def _abstract(blocks, spec, rep: Report):
    ab = spec["abstract"]
    abs_blocks = [b for b in blocks if b.role == M.ABSTRACT]
    if ab.get("structured") and abs_blocks and not any(b.label for b in abs_blocks):
        rep.warn("Format needs a structured abstract (" + " / ".join(ab.get("structured_labels", [])) +
                 ") but the manuscript abstract is unstructured — split it into those headings.")


def _equations(blocks, rep):
    eqs = [b for b in blocks if b.role == M.EQUATION]
    if any(b.eq_num for b in eqs):
        for i, b in enumerate(eqs, 1):
            b.eq_num = str(i)


# ------------------------------------------------------------------ references
def _references(blocks: List[M.Block], spec, rep: Report, crossref: bool, progress=None):
    style = spec["references"]["style"]
    mode = spec["citations"]["mode"]
    ref_blocks = [b for b in blocks if b.role == M.REFERENCE]
    if not ref_blocks:
        rep.warn("No reference list detected — citations were left unchanged.")
        return blocks
    refs = [parse_reference(b.text) for b in ref_blocks]
    if crossref:
        ok = 0
        for k, r in enumerate(refs):
            if progress:
                progress(k / max(1, len(refs)), f"Checking reference {k + 1}/{len(refs)} on Crossref")
            if enrich_crossref(r):
                ok += 1
        rep.change(f"{ok}/{len(refs)} references verified/completed via Crossref")
    bad = [r for r in refs if not r.confident]
    for r in bad:
        rep.warn("Reference could not be fully parsed, kept as written: " + r.raw[:90] + ("…" if len(r.raw) > 90 else ""))

    src_mode = detect_citation_mode(blocks)
    numeric_target = mode in ("numeric", "superscript")
    eng = CitationEngine(refs, mode, style)
    if src_mode == "author-year" and numeric_target:
        rep.change("In-text citations converted from author–year to numbered (order of first citation)")
    elif src_mode in ("numeric", "superscript") and not numeric_target:
        rep.change("In-text citations converted from numbered to author–year")
        rep.warn("Numbered citations were converted to author–year; check sentences where a number was used as a noun (e.g. “in [3]”).")
    elif src_mode != "none":
        rep.change(f"In-text citations re-styled ({mode})")
    for b in blocks:
        if b.role in (M.PARA, M.LIST, M.CAPTION, M.ABSTRACT):
            b.runs = eng.convert_runs(b.runs, src_mode)
    if eng.unresolved:
        uniq = list(dict.fromkeys(eng.unresolved))
        rep.warn(f"{len(uniq)} in-text citation(s) could not be matched to the reference list: " +
                 "; ".join(uniq[:8]) + (" …" if len(uniq) > 8 else ""))
    if numeric_target:
        cited = set(eng.order)
        newnum = eng.finalize_numbers(blocks)
        order = eng.order
        uncited = [k for k in range(len(refs)) if k not in cited]
        if uncited:
            rep.warn(f"{len(uncited)} reference(s) are never cited in the text (placed at the end of the list).")
    else:
        order = sort_author_year(refs)
        cited = set(eng.order)
        uncited = [k for k in range(len(refs)) if k not in cited]
        if src_mode != "none" and uncited:
            rep.warn(f"{len(uncited)} reference(s) are never cited in the text.")
    # rebuild the reference list
    rep.refs = [refs[k] for k in order]
    new_refs = []
    for n, k in enumerate(order, 1):
        b = M.Block(M.REFERENCE, runs=format_reference(refs[k], style, n))
        new_refs.append(b)
    rep.change(f"{len(refs)} references re-formatted to {style} style" +
               (" and ordered by first citation" if numeric_target else " and sorted alphabetically"))
    out = []
    inserted = False
    for b in blocks:
        if b.role == M.REFERENCE:
            if not inserted:
                out.extend(new_refs)
                inserted = True
            continue
        out.append(b)
    return out


def _ensure_ref_heading(blocks, spec):
    """Make sure the reference list is preceded by a heading with the format's wording."""
    head = spec["references"].get("heading", "References")
    for i, b in enumerate(blocks):
        if b.role == M.REFERENCE:
            if i > 0 and blocks[i - 1].role == M.HEADING:
                blocks[i - 1].runs = [Run(head)]
                blocks[i - 1].level = 1
            else:
                blocks.insert(i, M.Block(M.HEADING, runs=[Run(head)], level=1))
            return blocks
    return blocks


# ------------------------------------------------------------------ entry
def prepare(doc: M.Document, spec: dict, crossref: bool = False, progress=None) -> Tuple[M.Document, Report]:
    rep = Report()
    blocks = [copy.copy(b) for b in doc.blocks if b.role != M.ABSTRACT_HEADING]
    for w in doc.warnings:
        rep.warn(w)
    blocks = _pair_floats(blocks, spec, rep)
    maps = _number_captions(blocks, spec, rep)
    _fix_xrefs(blocks, spec, maps, rep)
    blocks = _references(blocks, spec, rep, crossref, progress)
    blocks = _ensure_ref_heading(blocks, spec)
    _number_headings(blocks, spec, rep)
    _keywords(blocks, spec, rep)
    _title(blocks, spec)
    _abstract(blocks, spec, rep)
    _equations(blocks, rep)
    for b in blocks:
        b.runs = merge_adjacent(b.runs)
    out = M.Document(blocks=blocks, source_type=doc.source_type, warnings=[])
    from .checks import run_checks
    run_checks(out, spec, rep)
    return out, rep
