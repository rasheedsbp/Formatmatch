"""Format specification: defaults, publisher presets, merging, journal lookup."""
from __future__ import annotations

import copy
import difflib
import json
import re
from typing import Optional

# ---------------------------------------------------------------- defaults
DEFAULT_SPEC = {
    "name": "Generic",
    "page": {"size": "A4", "margins_cm": [2.54, 2.54, 2.54, 2.54],  # top, bottom, left, right
             "columns": 1, "col_gap_cm": 0.5, "front_matter_single_column": True,
             "page_numbers": True, "line_numbers": False},
    "font": {"family": "Times New Roman", "size": 12, "line_spacing": 1.5, "align": "justify",
             "first_line_indent_cm": 0.0, "space_before_pt": 0, "space_after_pt": 6},
    "title": {"size": 16, "bold": True, "italic": False, "align": "center", "case": "as-is", "space_after_pt": 12},
    "authors": {"size": 11, "bold": False, "italic": False, "align": "center"},
    "affiliation": {"size": 10, "bold": False, "italic": True, "align": "center"},
    "abstract": {"heading": "Abstract", "layout": "block",          # block | inline
                 "suffix": ": ", "size": 11, "label_bold": True, "label_italic": False,
                 "body_bold": False, "body_italic": False, "align": "justify",
                 "structured": False,
                 "structured_labels": ["Purpose", "Design/methodology/approach", "Findings",
                                       "Originality/value"],
                 "indent_cm": 0.0},
    "keywords": {"label": "Keywords", "suffix": ": ", "separator": "; ", "size": 11,
                 "label_bold": True, "label_italic": False, "body_italic": False, "case": "as-is",
                 "end_period": False},
    "headings": {
        "scheme": "decimal",          # decimal | ieee | none
        "decimal_trailing_dot": True,  # "1. Introduction" vs "1 Introduction"
        "subsection_trailing_dot": True,  # "1.1. Data" vs "1.1 Data"
        "levels": [
            {"size": 12, "bold": True, "italic": False, "case": "as-is", "align": "left", "smallcaps": False,
             "space_before_pt": 12, "space_after_pt": 6},
            {"size": 12, "bold": True, "italic": False, "case": "as-is", "align": "left", "smallcaps": False,
             "space_before_pt": 10, "space_after_pt": 4},
            {"size": 12, "bold": True, "italic": True, "case": "as-is", "align": "left", "smallcaps": False,
             "space_before_pt": 8, "space_after_pt": 4},
            {"size": 12, "bold": False, "italic": True, "case": "as-is", "align": "left", "smallcaps": False,
             "space_before_pt": 6, "space_after_pt": 2},
        ],
    },
    "captions": {"figure_label": "Figure", "table_label": "Table", "table_numbering": "arabic",
                 "table_label_case": "as-is", "label_bold": True, "label_italic": False,
                 "separator": ". ", "size": 10, "align": "center", "table_position": "above",
                 "table_label_newline": False, "body_italic": False, "end_period": True},
    "tables": {"size": 10, "borders": "booktabs", "header_bold": True},
    "references": {"style": "apa", "heading": "References", "size": 11, "hanging_indent_cm": 1.27},
    "citations": {"mode": "author-year"},   # numeric | superscript | author-year
    "body_xref": {"figure": "Figure", "table": "Table"},
    "equations": {"align": "center"},
    "latex": {"class": "article", "options": "11pt,a4paper", "bibstyle": "", "packages": []},
    "limits": {"abstract_words": None, "keywords_min": None, "keywords_max": None,
               "title_words": None, "total_words": None},
    "required_sections": [],
}


def deep_merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if v is None:
            continue
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        elif k == "levels" and isinstance(v, list) and isinstance(out.get(k), list):
            merged = []
            for i in range(max(len(v), len(out[k]))):
                b = out[k][i] if i < len(out[k]) else {}
                o = v[i] if i < len(v) else {}
                merged.append(deep_merge(b, o or {}))
            out[k] = merged
        else:
            out[k] = copy.deepcopy(v)
    return out


def _lv(**kw):
    return kw


# ---------------------------------------------------------------- presets
PRESETS = {
    "ieee_journal": {
        "name": "IEEE Transactions / Journals (incl. IEEE Access)",
        "page": {"size": "Letter", "margins_cm": [1.9, 2.54, 1.57, 1.57], "columns": 2, "col_gap_cm": 0.43},
        "font": {"family": "Times New Roman", "size": 10, "line_spacing": 1.0, "first_line_indent_cm": 0.36,
                 "space_after_pt": 0},
        "title": {"size": 24, "bold": False, "case": "title"},
        "authors": {"size": 11},
        "affiliation": {"size": 8, "italic": False},
        "abstract": {"heading": "Abstract", "layout": "inline", "suffix": "—", "size": 9, "label_bold": True,
                     "label_italic": True, "body_bold": True},
        "keywords": {"label": "Index Terms", "suffix": "—", "separator": ", ", "size": 9, "label_bold": True,
                     "label_italic": True, "body_italic": False, "case": "lower-first", "end_period": True},
        "headings": {"scheme": "ieee", "levels": [
            _lv(size=10, bold=False, italic=False, case="title", align="center", smallcaps=True),
            _lv(size=10, bold=False, italic=True, case="title", align="left"),
            _lv(size=10, bold=False, italic=True, case="sentence", align="left"),
            _lv(size=10, bold=False, italic=True, case="sentence", align="left")]},
        "captions": {"figure_label": "Fig.", "table_label": "Table", "table_numbering": "roman",
                     "table_label_case": "upper", "label_bold": False, "separator": ". ", "size": 8,
                     "table_label_newline": True, "table_caption_case": "upper", "table_end_period": False},
        "tables": {"size": 8},
        "references": {"style": "ieee", "heading": "References", "size": 8, "hanging_indent_cm": 0.63},
        "citations": {"mode": "numeric"},
        "body_xref": {"figure": "Fig.", "table": "Table"},
        "latex": {"class": "IEEEtran", "options": "journal", "bibstyle": "IEEEtran"},
        "limits": {"abstract_words": 250, "keywords_min": 3, "keywords_max": 10},
    },
    "ieee_conference": {
        "name": "IEEE Conference",
        "page": {"size": "Letter", "margins_cm": [1.9, 2.54, 1.78, 1.78], "columns": 2, "col_gap_cm": 0.64},
        "font": {"family": "Times New Roman", "size": 10, "line_spacing": 1.0, "first_line_indent_cm": 0.36,
                 "space_after_pt": 0},
        "title": {"size": 24, "bold": False, "case": "title"},
        "authors": {"size": 11},
        "affiliation": {"size": 10, "italic": True},
        "abstract": {"heading": "Abstract", "layout": "inline", "suffix": "—", "size": 9, "label_bold": True,
                     "label_italic": True, "body_bold": True},
        "keywords": {"label": "Index Terms", "suffix": "—", "separator": ", ", "size": 9, "label_bold": True,
                     "label_italic": True, "body_bold": True, "case": "lower-first", "end_period": False},
        "headings": {"scheme": "ieee", "levels": [
            _lv(size=10, bold=False, italic=False, case="title", align="center", smallcaps=True),
            _lv(size=10, bold=False, italic=True, case="title", align="left"),
            _lv(size=10, bold=False, italic=True, case="sentence", align="left"),
            _lv(size=10, bold=False, italic=True, case="sentence", align="left")]},
        "captions": {"figure_label": "Fig.", "table_label": "Table", "table_numbering": "roman",
                     "table_label_case": "upper", "label_bold": False, "separator": ". ", "size": 8,
                     "table_label_newline": True, "table_caption_case": "upper", "table_end_period": False},
        "tables": {"size": 8},
        "references": {"style": "ieee", "heading": "References", "size": 8, "hanging_indent_cm": 0.63},
        "citations": {"mode": "numeric"},
        "body_xref": {"figure": "Fig.", "table": "Table"},
        "latex": {"class": "IEEEtran", "options": "conference", "bibstyle": "IEEEtran"},
        "limits": {"abstract_words": 200},
    },
    "elsevier": {
        "name": "Elsevier (elsarticle, numbered references)",
        "page": {"size": "A4", "margins_cm": [2.5, 2.5, 2.5, 2.5], "columns": 1, "line_numbers": True},
        "font": {"family": "Times New Roman", "size": 12, "line_spacing": 2.0},
        "title": {"size": 17, "bold": False, "case": "as-is"},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 11},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": "; ", "size": 11},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": True, "subsection_trailing_dot": True,
                     "levels": [_lv(size=12, bold=True), _lv(size=12, bold=False, italic=True),
                                _lv(size=12, bold=False, italic=True), _lv(size=12, bold=False, italic=True)]},
        "captions": {"figure_label": "Fig.", "table_label": "Table", "label_bold": True, "separator": ". ",
                     "size": 10},
        "references": {"style": "elsevier-numeric", "size": 11, "hanging_indent_cm": 0.75},
        "citations": {"mode": "numeric"},
        "body_xref": {"figure": "Fig.", "table": "Table"},
        "latex": {"class": "elsarticle", "options": "preprint,12pt", "bibstyle": "elsarticle-num"},
        "limits": {"abstract_words": 250, "keywords_min": 1, "keywords_max": 7},
    },
    "elsevier_harvard": {
        "name": "Elsevier (author–year / Harvard)",
        "_base": "elsevier",
        "references": {"style": "harvard"},
        "citations": {"mode": "author-year"},
        "latex": {"bibstyle": "elsarticle-harv"},
    },
    "springer_journal": {
        "name": "Springer Nature journals (sn-jnl)",
        "page": {"size": "A4", "margins_cm": [2.5, 2.5, 2.5, 2.5], "columns": 1},
        "font": {"family": "Times New Roman", "size": 11, "line_spacing": 1.5},
        "title": {"size": 16, "bold": True},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 10},
        "keywords": {"label": "Keywords", "suffix": " ", "separator": " · ", "size": 10},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": False, "subsection_trailing_dot": False,
                     "levels": [_lv(size=12, bold=True), _lv(size=11, bold=True),
                                _lv(size=11, bold=False, italic=True), _lv(size=11, italic=True, bold=False)]},
        "captions": {"figure_label": "Fig.", "table_label": "Table", "label_bold": True, "separator": " ",
                     "size": 9, "end_period": False},
        "references": {"style": "springer-basic", "size": 10},
        "citations": {"mode": "author-year-nocomma"},
        "body_xref": {"figure": "Fig.", "table": "Table"},
        "latex": {"class": "sn-jnl", "options": "pdflatex,sn-basic", "bibstyle": "sn-basic"},
        "limits": {"abstract_words": 250, "keywords_min": 4, "keywords_max": 6},
    },
    "springer_numeric": {
        "name": "Springer Nature journals (numbered references)",
        "_base": "springer_journal",
        "references": {"style": "springer-vancouver"},
        "citations": {"mode": "numeric"},
        "latex": {"options": "pdflatex,sn-mathphys-num", "bibstyle": "sn-mathphys-num"},
    },
    "springer_lncs": {
        "name": "Springer LNCS / conference proceedings",
        "page": {"size": "A4", "margins_cm": [5.25, 5.25, 4.4, 4.4], "columns": 1},
        "font": {"family": "Times New Roman", "size": 10, "line_spacing": 1.0, "first_line_indent_cm": 0.42,
                 "space_after_pt": 0},
        "title": {"size": 14, "bold": True},
        "authors": {"size": 10},
        "affiliation": {"size": 9, "italic": False},
        "abstract": {"heading": "Abstract", "layout": "inline", "suffix": ". ", "size": 9, "label_bold": True},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": " · ", "size": 9, "label_bold": True},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": False, "subsection_trailing_dot": False,
                     "levels": [_lv(size=12, bold=True), _lv(size=10, bold=True),
                                _lv(size=10, bold=True), _lv(size=10, italic=True, bold=False)]},
        "captions": {"figure_label": "Fig.", "table_label": "Table", "label_bold": True, "separator": ". ",
                     "size": 9},
        "tables": {"size": 9},
        "references": {"style": "lncs", "size": 9, "hanging_indent_cm": 0.5},
        "citations": {"mode": "numeric"},
        "body_xref": {"figure": "Fig.", "table": "Table"},
        "latex": {"class": "llncs", "options": "runningheads", "bibstyle": "splncs04"},
        "limits": {"abstract_words": 250, "keywords_min": 3, "keywords_max": 6},
    },
    "mdpi": {
        "name": "MDPI journals",
        "page": {"size": "A4", "margins_cm": [2.5, 2.5, 2.0, 2.0], "columns": 1},
        "font": {"family": "Palatino Linotype", "size": 10, "line_spacing": 1.0, "first_line_indent_cm": 0.75,
                 "space_after_pt": 0},
        "title": {"size": 18, "bold": True, "align": "left"},
        "authors": {"size": 10, "bold": True, "align": "left"},
        "affiliation": {"size": 8, "italic": False, "align": "left"},
        "abstract": {"heading": "Abstract", "layout": "inline", "suffix": ": ", "size": 9, "label_bold": True},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": "; ", "size": 9, "label_bold": True},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": True, "subsection_trailing_dot": True,
                     "levels": [_lv(size=10, bold=True), _lv(size=10, bold=False, italic=True),
                                _lv(size=10, bold=False, italic=False), _lv(size=10, italic=True, bold=False)]},
        "captions": {"figure_label": "Figure", "table_label": "Table", "label_bold": True, "separator": ". ",
                     "size": 9, "align": "left"},
        "tables": {"size": 9},
        "references": {"style": "mdpi", "size": 9, "hanging_indent_cm": 0.6},
        "citations": {"mode": "numeric"},
        "latex": {"class": "article", "options": "10pt,a4paper", "bibstyle": "mdpi"},
        "limits": {"abstract_words": 200, "keywords_min": 3, "keywords_max": 10},
    },
    "emerald": {
        "name": "Emerald journals (structured abstract, Harvard)",
        "page": {"size": "A4", "margins_cm": [2.54, 2.54, 2.54, 2.54], "columns": 1},
        "font": {"family": "Times New Roman", "size": 12, "line_spacing": 2.0},
        "abstract": {"heading": "Abstract", "layout": "block", "structured": True, "size": 12},
        "keywords": {"label": "Keywords", "suffix": " ", "separator": ", ", "size": 12, "case": "title"},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": True, "subsection_trailing_dot": False,
                     "levels": [_lv(size=12, bold=True), _lv(size=12, bold=True, italic=True),
                                _lv(size=12, bold=False, italic=True), _lv(size=12, italic=True, bold=False)]},
        "captions": {"figure_label": "Figure", "table_label": "Table", "label_bold": True, "separator": ". ",
                     "size": 11, "end_period": False},
        "references": {"style": "emerald-harvard", "size": 12},
        "citations": {"mode": "author-year"},
        "latex": {"class": "article", "options": "12pt,a4paper"},
        "limits": {"abstract_words": 250, "keywords_max": 12, "title_words": 15},
    },
    "taylor_francis": {
        "name": "Taylor & Francis (APA style)",
        "page": {"size": "A4", "margins_cm": [2.54, 2.54, 2.54, 2.54], "columns": 1},
        "font": {"family": "Times New Roman", "size": 12, "line_spacing": 2.0},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 12},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": "; ", "size": 12, "label_italic": False},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": True, "subsection_trailing_dot": False,
                     "levels": [_lv(size=12, bold=True), _lv(size=12, bold=True, italic=True),
                                _lv(size=12, bold=False, italic=True), _lv(size=12, italic=True, bold=False)]},
        "captions": {"figure_label": "Figure", "table_label": "Table", "label_bold": False, "separator": ". ",
                     "size": 11},
        "references": {"style": "apa", "size": 12},
        "citations": {"mode": "author-year"},
        "latex": {"class": "article", "options": "12pt,a4paper"},
        "limits": {"abstract_words": 200, "keywords_min": 3, "keywords_max": 6},
    },
    "wiley": {
        "name": "Wiley journals (generic, APA)",
        "_base": "taylor_francis",
        "references": {"style": "apa"},
    },
    "apa7": {
        "name": "APA 7th edition (theses, APA journals)",
        "page": {"size": "Letter", "margins_cm": [2.54, 2.54, 2.54, 2.54], "columns": 1},
        "font": {"family": "Times New Roman", "size": 12, "line_spacing": 2.0, "first_line_indent_cm": 1.27,
                 "space_after_pt": 0, "align": "left"},
        "title": {"size": 12, "bold": True, "case": "title"},
        "authors": {"size": 12}, "affiliation": {"size": 12, "italic": False},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 12, "align": "left"},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": ", ", "size": 12, "label_italic": True,
                     "label_bold": False},
        "headings": {"scheme": "none", "levels": [
            _lv(size=12, bold=True, case="title", align="center"),
            _lv(size=12, bold=True, case="title", align="left"),
            _lv(size=12, bold=True, italic=True, case="title", align="left"),
            _lv(size=12, bold=True, case="title", align="left")]},
        "captions": {"figure_label": "Figure", "table_label": "Table", "label_bold": True, "separator": "\n",
                     "size": 12, "align": "left", "body_italic": True, "end_period": False,
                     "figure_position": "above"},
        "references": {"style": "apa", "size": 12, "hanging_indent_cm": 1.27},
        "citations": {"mode": "author-year"},
        "latex": {"class": "article", "options": "12pt,letterpaper"},
        "limits": {"abstract_words": 250, "keywords_min": 3, "keywords_max": 5},
    },
    "plos": {
        "name": "PLOS (Vancouver)",
        "page": {"size": "Letter", "columns": 1, "line_numbers": True},
        "font": {"family": "Arial", "size": 10, "line_spacing": 1.5, "align": "left"},
        "title": {"size": 18, "bold": True, "align": "left"},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 10},
        "headings": {"scheme": "none", "levels": [_lv(size=14, bold=True), _lv(size=12, bold=True),
                                                  _lv(size=10, bold=True), _lv(size=10, italic=True, bold=False)]},
        "captions": {"figure_label": "Fig", "table_label": "Table", "separator": ". ", "size": 10, "align": "left"},
        "references": {"style": "vancouver", "size": 10, "hanging_indent_cm": 0.63},
        "citations": {"mode": "numeric"},
        "body_xref": {"figure": "Fig", "table": "Table"},
        "latex": {"class": "article", "options": "10pt,letterpaper"},
        "limits": {"abstract_words": 300},
    },
    "frontiers": {
        "name": "Frontiers journals",
        "page": {"size": "A4", "columns": 1},
        "font": {"family": "Times New Roman", "size": 12, "line_spacing": 1.15},
        "abstract": {"heading": "Abstract", "layout": "block", "size": 12},
        "keywords": {"label": "Keywords", "suffix": ": ", "separator": ", "},
        "headings": {"scheme": "decimal", "decimal_trailing_dot": False, "subsection_trailing_dot": False,
                     "levels": [_lv(bold=True, case="upper"), _lv(bold=True), _lv(bold=True, italic=True),
                                _lv(italic=True, bold=False)]},
        "captions": {"figure_label": "Figure", "table_label": "Table", "label_bold": True, "separator": " "},
        "references": {"style": "harvard", "size": 12},
        "citations": {"mode": "author-year"},
        "limits": {"abstract_words": 350, "keywords_min": 5, "keywords_max": 8},
    },
    "generic_numeric": {
        "name": "Generic numbered (Vancouver)",
        "references": {"style": "vancouver"}, "citations": {"mode": "numeric"},
    },
}


def get_preset(key: str) -> dict:
    p = copy.deepcopy(PRESETS[key])
    base = p.pop("_base", None)
    spec = get_preset(base) if base else copy.deepcopy(DEFAULT_SPEC)
    return deep_merge(spec, p)


# ---------------------------------------------------------------- journal name → preset
PUBLISHER_MAP = [
    (r"\bieee\b|institute of electrical", "ieee_journal"),
    (r"elsevier|cell press|academic press", "elsevier"),
    (r"lecture notes|lncs|ccis", "springer_lncs"),
    (r"springer|nature|bmc|biomed central|palgrave", "springer_numeric"),
    (r"mdpi|multidisciplinary digital", "mdpi"),
    (r"emerald", "emerald"),
    (r"taylor|francis|informa|routledge", "taylor_francis"),
    (r"wiley|hindawi|blackwell", "wiley"),
    (r"public library of science|\bplos\b", "plos"),
    (r"frontiers", "frontiers"),
    (r"american psychological|\bapa\b", "apa7"),
]
KNOWN_JOURNALS = {
    "ieee access": "ieee_journal", "international journal of organizational analysis": "emerald",
    "expert systems with applications": "elsevier", "computers & security": "elsevier",
    "computers and security": "elsevier", "future generation computer systems": "elsevier",
    "applied energy": "elsevier", "journal of energy storage": "elsevier",
    "scientific reports": "springer_numeric", "sensors": "mdpi", "electronics": "mdpi",
    "sustainability": "mdpi", "applied sciences": "mdpi", "energies": "mdpi",
    "plos one": "plos", "management decision": "emerald", "journal of knowledge management": "emerald",
    "international journal of intelligent engineering and systems": "ieee_journal",
}


def preset_for_publisher(text: str) -> Optional[str]:
    t = (text or "").lower()
    for pat, key in PUBLISHER_MAP:
        if re.search(pat, t):
            if key == "ieee_journal" and re.search(r"conference|proceedings|symposium|workshop", t):
                return "ieee_conference"
            return key
    return None


def lookup_journal(name: str, timeout: float = 8.0) -> dict:
    """Resolve a journal name to {title, publisher, issn, preset, source}. Network optional."""
    n = (name or "").strip()
    res = {"query": n, "title": n, "publisher": "", "issn": "", "preset": None, "source": ""}
    if not n:
        return res
    key = n.lower()
    if key in KNOWN_JOURNALS:
        res.update(preset=KNOWN_JOURNALS[key], source="built-in list")
    direct = preset_for_publisher(n)
    if not res["preset"] and direct:
        res.update(preset=direct, source="journal name")
    try:
        import requests
        r = requests.get("https://api.crossref.org/journals", params={"query": n, "rows": 8},
                         headers={"User-Agent": "FormatMatch/1.0 (mailto:info@cssoftwaresolutions.tech)"},
                         timeout=timeout)
        items = r.json()["message"]["items"]
        if items:
            best = max(items, key=lambda it: difflib.SequenceMatcher(None, it.get("title", "").lower(), key).ratio())
            score = difflib.SequenceMatcher(None, best.get("title", "").lower(), key).ratio()
            if score > 0.6:
                res["title"] = best.get("title", n)
                res["publisher"] = best.get("publisher", "")
                res["issn"] = ", ".join(best.get("ISSN", []))
                pp = preset_for_publisher(res["publisher"] + " " + res["title"])
                if pp and not res["preset"]:
                    res.update(preset=pp, source="Crossref publisher")
    except Exception:
        pass
    return res


# ---------------------------------------------------------------- guidelines text → overrides
FONT_NAMES = ["Times New Roman", "Arial", "Calibri", "Cambria", "Palatino Linotype", "Palatino", "Garamond",
              "Helvetica", "Georgia", "Book Antiqua", "Verdana", "Computer Modern"]
REF_STYLE_WORDS = [
    (r"\bieee\b(?:\s+(?:reference|citation|style))", "ieee"), (r"\bapa\b", "apa"), (r"harvard", "harvard"),
    (r"vancouver", "vancouver"), (r"chicago", "chicago"), (r"\bama\b|american medical association", "ama"),
    (r"\bmla\b", "mla"), (r"springer basic|sn-basic", "springer-basic"),
]
NUM_WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8,
             "nine": 9, "ten": 10, "twelve": 12, "fifteen": 15}


def _num(s):
    s = s.lower()
    return NUM_WORDS.get(s) or (float(s) if re.fullmatch(r"\d+(?:\.\d+)?", s) else None)


def parse_guidelines(text: str) -> dict:
    """Rule-based extraction of formatting requirements from author-guidelines text."""
    t = re.sub(r"\s+", " ", text or "")
    tl = t.lower()
    o: dict = {"page": {}, "font": {}, "abstract": {}, "keywords": {}, "references": {}, "citations": {},
               "limits": {}, "headings": {}}
    found = []

    for f in FONT_NAMES:
        if f.lower() in tl:
            o["font"]["family"] = "Palatino Linotype" if f == "Palatino" else f
            found.append(f"font: {f}")
            break
    m = re.search(r"(\d{1,2}(?:\.\d)?)\s*-?\s*(?:pt|point)\b", tl)
    if m:
        o["font"]["size"] = float(m.group(1))
        found.append(f"font size: {m.group(1)} pt")
    if re.search(r"double[- ]spac", tl):
        o["font"]["line_spacing"] = 2.0; found.append("double spacing")
    elif re.search(r"1\.5[- ]?(?:line)?[- ]?spac|one and a half", tl):
        o["font"]["line_spacing"] = 1.5; found.append("1.5 spacing")
    elif re.search(r"single[- ]spac", tl):
        o["font"]["line_spacing"] = 1.0; found.append("single spacing")
    m = re.search(r"margins?[^.]{0,40}?(\d(?:\.\d+)?)\s*(cm|mm|inch|inches|in\b|\")", tl)
    if m:
        v = float(m.group(1)); u = m.group(2)
        cm = v * 2.54 if u.startswith("in") or u == '"' else (v / 10 if u == "mm" else v)
        o["page"]["margins_cm"] = [round(cm, 2)] * 4; found.append(f"margins: {v}{u}")
    if re.search(r"\ba4\b", tl):
        o["page"]["size"] = "A4"; found.append("A4 paper")
    elif re.search(r"\bus letter\b|letter[- ]size|8\.5\s*[x×]\s*11", tl):
        o["page"]["size"] = "Letter"; found.append("Letter paper")
    if re.search(r"two[- ]column|double[- ]column", tl):
        o["page"]["columns"] = 2; found.append("two columns")
    elif re.search(r"single[- ]column|one[- ]column", tl):
        o["page"]["columns"] = 1; found.append("single column")
    if re.search(r"line number", tl):
        o["page"]["line_numbers"] = True; found.append("line numbers")

    m = re.search(r"abstract[^.]{0,80}?(?:not exceed|maximum of|max(?:imum)?\.?|no more than|up to|within|limited to|of about|of)\s*(\d{2,3})\s*words", tl) \
        or re.search(r"abstract[^.]{0,30}?\(?(\d{2,3})\s*words", tl) \
        or re.search(r"(\d{2,3})[- ]word abstract", tl)
    if m:
        o["limits"]["abstract_words"] = int(m.group(1)); found.append(f"abstract ≤ {m.group(1)} words")
    if re.search(r"structured abstract", tl) or re.search(r"purpose.{0,40}design/methodology/approach", tl):
        o["abstract"]["structured"] = True; found.append("structured abstract")
    m = re.search(r"(\w+)\s*(?:to|-|–)\s*(\w+)\s*key\s*words", tl)
    if m and _num(m.group(1)) and _num(m.group(2)):
        o["limits"]["keywords_min"] = int(_num(m.group(1))); o["limits"]["keywords_max"] = int(_num(m.group(2)))
        found.append(f"{m.group(1)}–{m.group(2)} keywords")
    else:
        m = re.search(r"(?:up to|maximum of|max\.?|no more than)\s*(\w+)\s*key\s*words", tl)
        if m and _num(m.group(1)):
            o["limits"]["keywords_max"] = int(_num(m.group(1))); found.append(f"≤ {m.group(1)} keywords")
    m = re.search(r"title[^.]{0,40}?(?:not exceed|maximum of|no more than|up to)\s*(\d{1,3})\s*words", tl)
    if m:
        o["limits"]["title_words"] = int(m.group(1)); found.append(f"title ≤ {m.group(1)} words")
    m = re.search(r"(?:manuscript|paper|article|text)[^.]{0,60}?(?:not exceed|maximum of|no more than|up to|within)\s*([\d,]{4,6})\s*words", tl)
    if m:
        o["limits"]["total_words"] = int(m.group(1).replace(",", "")); found.append(f"≤ {m.group(1)} words total")

    for pat, style in REF_STYLE_WORDS:
        if re.search(pat, tl):
            o["references"]["style"] = style
            o["citations"]["mode"] = "numeric" if style in ("ieee", "vancouver", "ama") else "author-year"
            found.append(f"reference style: {style}")
            break
    if "style" not in o["references"] and re.search(r"numbered (?:consecutively )?in (?:the )?order|square brackets", tl):
        o["citations"]["mode"] = "numeric"; o["references"].setdefault("style", "vancouver")
        found.append("numbered citations")
    if re.search(r"(?:headings?|sections?)[^.]{0,40}(?:should not|must not|not) be numbered|unnumbered (?:sections|headings)", tl):
        o["headings"]["scheme"] = "none"; found.append("unnumbered headings")
    elif re.search(r"(?:headings?|sections?)[^.]{0,40}(?:should|must) be numbered|numbered (?:sections|headings)", tl):
        o["headings"]["scheme"] = "decimal"; found.append("numbered headings")

    req = []
    for sec in ["Introduction", "Literature Review", "Methodology", "Methods", "Results", "Discussion",
                "Conclusion", "Acknowledgements", "Data Availability", "Conflict of Interest", "Funding",
                "Author Contributions", "Declaration of Competing Interest", "Ethics Statement"]:
        if re.search(r"(?:must|should|required|include|mandatory)[^.]{0,120}\b" + sec.lower(), tl):
            req.append(sec)
    if req:
        o["required_sections"] = req; found.append("required sections: " + ", ".join(req))
    o = {k: v for k, v in o.items() if v}
    o["_found"] = found
    return o


def parse_guidelines_llm(text: str, api_key: str, model: str = "llama-3.3-70b-versatile") -> dict:
    """Optional: ask a Groq-hosted LLM to map guidelines to spec overrides (JSON)."""
    import requests
    schema_hint = json.dumps({k: DEFAULT_SPEC[k] for k in
                              ["page", "font", "title", "abstract", "keywords", "headings", "captions",
                               "references", "citations", "limits", "required_sections"]})
    prompt = ("Extract manuscript formatting rules from the author guidelines below. Return ONLY a JSON object "
              "with the same keys/structure as this schema, filling only values the guidelines state "
              "(omit unknowns). references.style must be one of: ieee, apa, harvard, emerald-harvard, vancouver, "
              "chicago, ama, mla, springer-basic, springer-vancouver, lncs, mdpi, elsevier-numeric. "
              "citations.mode one of numeric, superscript, author-year. headings.scheme one of decimal, ieee, none."
              f"\n\nSCHEMA:\n{schema_hint}\n\nGUIDELINES:\n{text[:24000]}")
    r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                      headers={"Authorization": f"Bearer {api_key}"},
                      json={"model": model, "temperature": 0, "response_format": {"type": "json_object"},
                            "messages": [{"role": "user", "content": prompt}]}, timeout=90)
    r.raise_for_status()
    data = json.loads(r.json()["choices"][0]["message"]["content"])
    data["_found"] = ["LLM-extracted rules (review below)"]
    return data
