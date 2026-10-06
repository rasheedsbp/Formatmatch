"""Role classification shared by the DOCX/PDF readers and the template analyser."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from . import model as M

KNOWN_SECTIONS = [
    "introduction", "related work", "related works", "literature review", "review of literature",
    "background", "preliminaries", "methodology", "methods", "method", "materials and methods",
    "proposed method", "proposed methodology", "proposed model", "proposed system", "proposed framework",
    "proposed approach", "system model", "problem formulation", "research methodology",
    "experimental setup", "experiments", "experimental results", "results", "discussion",
    "results and discussion", "results and discussions", "performance evaluation", "analysis",
    "conclusion", "conclusions", "conclusion and future work", "conclusions and future work",
    "future work", "limitations", "implications", "theoretical implications", "practical implications",
    "acknowledgment", "acknowledgments", "acknowledgement", "acknowledgements",
    "references", "bibliography", "appendix", "appendices", "data availability",
    "data availability statement", "conflict of interest", "conflicts of interest",
    "declaration of competing interest", "funding", "author contributions", "credit authorship contribution statement",
    "ethics statement", "nomenclature", "abbreviations", "hypotheses", "hypotheses development",
    "theoretical framework", "theoretical background", "conceptual framework", "case study",
    "dataset", "datasets", "evaluation", "simulation results", "findings",
]
UNNUMBERED = {
    "acknowledgment", "acknowledgments", "acknowledgement", "acknowledgements", "references",
    "bibliography", "data availability", "data availability statement", "conflict of interest",
    "conflicts of interest", "declaration of competing interest", "funding", "author contributions",
    "credit authorship contribution statement", "ethics statement", "nomenclature", "abbreviations",
    "appendix", "appendices",
}
REF_HEADINGS = {"references", "bibliography", "literature cited", "works cited", "reference", "list of references"}

ROMAN = r"(?:X{0,3})(?:IX|IV|V?I{0,3})"
NUM_RE = re.compile(
    r"^\s*(?:"
    r"(?P<num>\d{1,2}(?:\.\d{1,2}){0,3})\.?\)?"
    r"|(?P<roman>" + ROMAN + r")\."
    r"|(?P<alpha>[A-H])\."
    r"|(?P<paren>\d{1,2})\)"
    r")\s+(?P<rest>\S.*)$"
)
CAPTION_RE = re.compile(
    r"^\s*(?P<kind>fig\.?|figure|table|tab\.)\s*(?P<num>\d+|[IVXLC]+)\s*(?P<sep>[\.:\-–—|]\s*)?(?P<rest>.*)$",
    re.I | re.S,
)
CAPTION_VERBS = {
    "shows", "show", "illustrates", "presents", "depicts", "depict", "demonstrates", "compares", "summarizes",
    "summarises", "lists", "gives", "reports", "displays", "indicates", "contains", "provides", "describes",
    "reveals", "plots", "highlights", "outlines", "is", "are", "was", "were", "and", "also", "in", "of", "confirms",
    "represents", "exhibits", "details", "clearly", "further", "below", "above", "visualizes", "visualises",
    "tabulates", "explains", "suggests", "has", "have", "can", "will", "we", "it", "the", "a", "an", "to",
}
ARTICLE_TYPE_RE = re.compile(
    r"^\s*(?:(?:original|research|review|short|brief|technical|case|mini|systematic|full[- ]length|regular|"
    r"invited|perspective|empirical|conceptual|methodological)\s+)*(?:research\s+)?(?:article|paper|communication|"
    r"report|note|review|letter|study|manuscript)s?\s*$|"
    r"^\s*(?:running\s+(?:head|title)|short\s+title|article\s+type|manuscript\s+(?:type|id|no|number)|"
    r"paper\s+type|received|accepted|published|available\s+online|doi|issn|e-?issn|vol\.|volume)\b",
    re.I,
)
TITLE_PREFIX_RE = re.compile(r"^\s*(?:paper\s+|article\s+|manuscript\s+)?title\s*[:\-–—]\s*(?P<rest>\S.*)$", re.I | re.S)


def is_caption_text(t: str, style: str = "") -> bool:
    """Is a paragraph a figure/table caption (and not a sentence that merely starts with 'Figure 2 shows…')?"""
    cm = CAPTION_RE.match(t)
    if not cm or len(t) > 600:
        return False
    if (style or "").lower().startswith("caption"):
        return True
    rest = cm.group("rest").strip()
    first = re.match(r"[A-Za-z]+", rest)
    first_word = first.group(0).lower() if first else ""
    if not rest:
        return len(t) < 24                                  # "Table 1" alone: text follows on the next line
    if cm.group("sep"):
        return first_word not in CAPTION_VERBS or rest[:1].isupper()
    if cm.group("kind").isupper() and len(cm.group("kind")) > 2:
        return True                                         # "TABLE 2 HYPER-PARAMETERS"
    # no separator: "Fig 2 Training loss …" → caption ; "Figure 2 shows …" → sentence
    return (rest[:1].isupper() or rest[:1] in "(“\"'") and first_word not in CAPTION_VERBS


ABSTRACT_RE = re.compile(r"^\s*abstract\s*(?:[:\-–—.]\s*)?(?P<rest>.*)$", re.I | re.S)
KEYWORDS_RE = re.compile(r"^\s*(?:key\s*-?\s*words?|index\s+terms)\s*(?:[:\-–—.]\s*)?(?P<rest>.*)$", re.I | re.S)
STRUCT_ABS_RE = re.compile(
    r"^\s*(?P<lab>purpose|design/methodology/approach|design|methodology|approach|findings|"
    r"research limitations/implications|research limitations|practical implications|social implications|"
    r"originality/value|originality|value|background|objective|objectives|aim|aims|methods|results|"
    r"conclusions?|significance)\s*[:\-–—]\s*(?P<rest>\S.*)$",
    re.I | re.S,
)
AFFIL_HINT = re.compile(
    r"universit|department|dept\.|institute|college|school of|faculty|laborator|centre|center for|"
    r"@|e-?mail|india|china|usa|u\.s\.a|united kingdom|\bUK\b|road|street|pin\b|zip|corresponding",
    re.I,
)

ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}


def roman_to_int(s: str) -> int:
    s = s.upper()
    total, prev = 0, 0
    for ch in reversed(s):
        v = ROMAN_VALUES.get(ch, 0)
        total = total - v if v < prev else total + v
        prev = max(prev, v)
    return total


def int_to_roman(n: int) -> str:
    vals = [(100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def strip_heading_number(text: str):
    """Return (level_from_numbering or 0, bare heading text)."""
    t = norm(text)
    m = NUM_RE.match(t)
    if not m:
        return 0, t
    rest = m.group("rest")
    if m.group("num"):
        return m.group("num").count(".") + 1, rest
    if m.group("roman") and m.group("roman") != "":
        return 1, rest
    if m.group("alpha"):
        return 2, rest
    if m.group("paren"):
        return 3, rest
    return 0, t


def known_section(text: str) -> Optional[str]:
    _, bare = strip_heading_number(text)
    key = re.sub(r"[^a-z/ ]", "", bare.lower()).strip()
    if key in KNOWN_SECTIONS or key in REF_HEADINGS:
        return key
    return None


def looks_like_heading(text: str, bold: bool, size: float, body_size: float, style: str) -> int:
    """Return heading level (1..4) or 0."""
    t = norm(text)
    if not t or len(t) > 160:
        return 0
    s = (style or "").lower()
    m = re.match(r"heading\s*(\d)", s)
    if m:
        return min(int(m.group(1)), 4)
    if s in ("title", "subtitle"):
        return 0
    level, bare = strip_heading_number(t)
    words = len(bare.split())
    # pseudo-code / step lines ("1. Initialize population", "2. x ← x + 1") are not headings
    from .algo import CODEISH_RE, STEP_RE
    if STEP_RE.match(bare) and not known_section(t) or CODEISH_RE.search(bare) or re.search(r"[=←∑∈≤≥]", bare):
        return 0
    ends_sentence = bare.endswith((".", ",", ";")) and not bare.endswith("etc.")
    ks = known_section(t)
    if ks and words <= 8:
        return level or 1
    if level and words <= 16 and not ends_sentence and (bold or (size and size > body_size + 0.4) or bare[:1].isupper()):
        # numbered short line: very likely a heading; guard against numbered list items
        if bold or (size and size > body_size + 0.4) or bare.isupper() or words <= 8:
            return level
    if bold and words <= 12 and not ends_sentence and len(t) < 100 and not CAPTION_RE.match(t):
        return 2 if not (size and size > body_size + 1) else 1
    return 0


@dataclass
class ParaInfo:
    text: str
    style: str = ""
    size: float = 0.0
    bold: bool = False
    is_list: bool = False
    list_ordered: bool = False
    list_level: int = 0
    align: str = ""
    is_title_style: bool = False


class RoleClassifier:
    """Streaming classifier: feed paragraphs in reading order."""

    def __init__(self, body_size: float = 0.0, max_size: float = 0.0):
        self.state = "front"   # front -> abstract -> body -> refs
        self.body_size = body_size
        self.max_size = max_size
        self.seen_title = False
        self.count = 0

    def feed(self, p: ParaInfo):
        """Return (role, level, label, cleaned_text_or_None)."""
        t = norm(p.text)
        self.count += 1
        style = (p.style or "").lower()

        if not t:
            return None, 0, "", None

        # captions are recognisable anywhere
        cm = CAPTION_RE.match(t)
        if cm and self.state not in ("refs", "front") and is_caption_text(t, style):
            return M.CAPTION, 0, cm.group("kind").lower(), None

        # explicit styles (Word templates / pandoc output)
        if style == "bibliography":
            self.state = "refs"
            return M.REFERENCE, 0, "", None
        if self.state in ("front", "abstract"):
            if style == "author":
                self.seen_title = True
                return M.AUTHOR, 0, "", None
            if style in ("date",):
                return M.OTHER_FRONT, 0, "", None
            if style in ("abstract title", "abstracttitle"):
                self.state = "abstract"
                return "abstract_label", 0, "", ""
            if style == "abstract":
                self.state = "abstract"
                am = ABSTRACT_RE.match(t)
                if am and t.lower().startswith("abstract"):
                    rest = am.group("rest").strip()
                    return (M.ABSTRACT, 0, "", rest) if rest else ("abstract_label", 0, "", "")
                km = KEYWORDS_RE.match(t)
                if km:
                    self.state = "body_pending"
                    return M.KEYWORDS, 0, "", km.group("rest").strip()
                return M.ABSTRACT, 0, "", None

        if self.state == "refs":
            lvl = looks_like_heading(t, p.bold, p.size, self.body_size, p.style)
            ks = known_section(t)
            if lvl and ks and ks not in REF_HEADINGS:
                self.state = "body"
                return M.HEADING, lvl, "", None
            if lvl and (t.lower().startswith("appendix") or style.startswith("heading")):
                self.state = "body"
                return M.HEADING, lvl, "", None
            return M.REFERENCE, 0, "", None

        # ---- front matter ----
        if self.state in ("front", "abstract"):
            am = ABSTRACT_RE.match(t)
            if am and (len(t) < 15 or am.group("rest")) and t.lower().startswith("abstract"):
                self.state = "abstract"
                rest = am.group("rest").strip()
                if not rest:
                    return "abstract_label", 0, "", ""
                return M.ABSTRACT, 0, "", rest
            km = KEYWORDS_RE.match(t)
            if km and len(t) < 700:
                self.state = "body_pending"
                return M.KEYWORDS, 0, "", km.group("rest").strip()
            if self.state == "abstract":
                lvl = looks_like_heading(t, p.bold, p.size, self.body_size, p.style)
                if lvl and known_section(t):
                    self.state = "body"
                    return M.HEADING, lvl, "", None
                sm = STRUCT_ABS_RE.match(t)
                if sm:
                    return M.ABSTRACT, 0, sm.group("lab").strip(), sm.group("rest").strip()
                return M.ABSTRACT, 0, "", None

            # still before the abstract
            lvl = looks_like_heading(t, p.bold, p.size, self.body_size, p.style)
            if lvl and known_section(t) and self.seen_title:
                self.state = "body"
                return M.HEADING, lvl, "", None
            if not self.seen_title:
                tm = TITLE_PREFIX_RE.match(t)
                if tm:
                    self.seen_title = True
                    return M.TITLE, 0, "", tm.group("rest").strip()
                # labels above the title: "Original Research Article", "Running head: …", journal banner lines
                if ARTICLE_TYPE_RE.match(t) and not p.is_title_style and style != "title":
                    return M.OTHER_FRONT, 0, "", None
                if (self.max_size >= self.body_size + 3 and p.size < self.max_size - 1.5 and len(t.split()) <= 6
                        and not p.is_title_style and style != "title"):
                    return M.OTHER_FRONT, 0, "", None
                if p.is_title_style or style == "title" or (
                    len(t) < 300 and (p.size >= max(self.body_size + 2, self.max_size - 0.5) or p.bold or self.count <= 2)
                ):
                    self.seen_title = True
                    return M.TITLE, 0, "", None
                # long first paragraph: no title detected -> treat as body start
                if len(t) > 300:
                    self.state = "body"
                    return M.PARA, 0, "", None
                self.seen_title = True
                return M.TITLE, 0, "", None
            if style == "subtitle":
                return M.OTHER_FRONT, 0, "", None
            if AFFIL_HINT.search(t):
                return M.AFFIL, 0, "", None
            if len(t) < 400 and not t.endswith(".") or re.search(r"\b(and|&)\b|,", t) and len(t) < 300:
                return M.AUTHOR, 0, "", None
            if len(t) > 400:
                self.state = "body"
                return M.PARA, 0, "", None
            return M.OTHER_FRONT, 0, "", None

        # ---- body ----
        if self.state == "body_pending":
            self.state = "body"
        lvl = looks_like_heading(t, p.bold, p.size, self.body_size, p.style)
        if lvl and not p.is_list:
            ks = known_section(t)
            if ks in REF_HEADINGS:
                self.state = "refs"
            return M.HEADING, lvl, "", None
        if p.is_list:
            return M.LIST, p.list_level, "", None
        return M.PARA, 0, "", None
