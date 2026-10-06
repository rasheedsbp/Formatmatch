"""Helpers for editing runs while keeping inline formatting."""
from __future__ import annotations

import copy
import re
from typing import List

from .model import Run


def trim_leading(runs: List[Run], rest: str) -> List[Run]:
    """Drop leading characters so the plain text starts with `rest` (whitespace-insensitive)."""
    if rest is None:
        return runs
    full = "".join(r.text for r in runs if not r.is_math)
    target = re.sub(r"\s+", " ", rest).strip()
    if not target:
        return []
    probe = target[: min(12, len(target))]
    norm_full = full
    idx = norm_full.find(probe)
    if idx < 0:
        # whitespace differences: search loosely
        pattern = r"\s*".join(re.escape(c) for c in probe if not c.isspace())
        m = re.search(pattern, norm_full)
        if not m:
            return [Run(target)]
        idx = m.start()
    return drop_chars(runs, idx)


def drop_chars(runs: List[Run], n: int) -> List[Run]:
    out: List[Run] = []
    remaining = n
    for r in runs:
        if remaining <= 0:
            out.append(r)
            continue
        if r.is_math:
            out.append(r)
            continue
        if len(r.text) <= remaining:
            remaining -= len(r.text)
            continue
        nr = copy.copy(r)
        nr.text = r.text[remaining:]
        remaining = 0
        out.append(nr)
    if out and not out[0].is_math:
        out[0] = copy.copy(out[0])
        out[0].text = out[0].text.lstrip()
    return [r for r in out if r.text or r.is_math]


def strip_trailing(runs: List[Run], chars: str = " \t") -> List[Run]:
    runs = [copy.copy(r) for r in runs]
    while runs and not runs[-1].is_math:
        runs[-1].text = runs[-1].text.rstrip(chars)
        if runs[-1].text:
            break
        runs.pop()
    return runs


def merge_adjacent(runs: List[Run]) -> List[Run]:
    out: List[Run] = []
    for r in runs:
        if (
            out
            and not r.is_math
            and not out[-1].is_math
            and not r.cite
            and not out[-1].cite
            and (r.bold, r.italic, r.sup, r.sub) == (out[-1].bold, out[-1].italic, out[-1].sup, out[-1].sub)
        ):
            out[-1] = copy.copy(out[-1])
            out[-1].text += r.text
        else:
            out.append(r)
    return out


def _cap(w: str) -> str:
    """Capitalise a word's first *Latin* letter only (keeps α-, β-, 3D, iPhone …)."""
    if not w:
        return w
    c = w[0]
    if c.isascii() or "\u00c0" <= c <= "\u024f":
        return c.upper() + w[1:]
    return w


def apply_case(text: str, case: str) -> str:
    if case in ("title", "sentence") and text.isupper() and len(text) > 4:
        text = text.lower()   # all-caps source heading: no acronym info to preserve
    if case == "upper":
        return text.upper()
    if case == "title":
        small = {"a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into", "nor", "of",
                 "on", "or", "over", "the", "to", "via", "vs", "with", "without"}
        words = text.split(" ")
        out = []
        for i, w in enumerate(words):
            if w.isupper() and len(w) > 1:          # acronyms
                out.append(w)
            elif i > 0 and w.lower() in small:
                out.append(w.lower())
            else:
                out.append(_cap(w[:1] + w[1:].lower()) if not any(c.isupper() for c in w[1:]) else _cap(w))
        return " ".join(out)
    if case == "sentence":
        words = text.split(" ")
        out = []
        for i, w in enumerate(words):
            if (w.isupper() and len(w) > 1) or any(c.isupper() for c in w[1:]):
                out.append(w)                         # keep acronyms / CamelCase
            elif i == 0:
                out.append(_cap(w[:1] + w[1:].lower()))
            else:
                out.append(w.lower())
        return " ".join(out)
    return text
