"""Learn a format from a template / sample paper (DOCX or PDF)."""
from __future__ import annotations

import collections
import copy
import io
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from docx.oxml.ns import qn
from lxml import etree

from . import model as M
from .classify import CAPTION_RE, KEYWORDS_RE, ABSTRACT_RE, NUM_RE, strip_heading_number
from .refs import BRACKET_RE, NARR_RE, PAREN_RE, parse_reference

TWIP_CM = 2.54 / 1440


@dataclass
class TemplateProfile:
    data: bytes
    exemplars: Dict[str, dict] = field(default_factory=dict)
    sections: List[str] = field(default_factory=list)       # sectPr XML strings, in order
    front_section: Optional[str] = None                     # sectPr closing the front matter (if any)
    table_tblPr: Optional[str] = None
    found: List[str] = field(default_factory=list)


# ------------------------------------------------------------------ helpers
def _x(el) -> Optional[str]:
    return etree.tostring(el, encoding="unicode") if el is not None else None


def _dominant_rpr(p_el, skip_label: bool = False):
    """rPr of the run carrying most text (optionally ignoring the first label run)."""
    best, best_len, first_rpr = None, -1, None
    runs = list(p_el.iter(qn("w:r")))
    for k, r in enumerate(runs):
        t = "".join(x.text or "" for x in r.iter(qn("w:t")))
        if not t.strip():
            continue
        if first_rpr is None:
            first_rpr = r.find(qn("w:rPr"))
            if skip_label:
                continue
        if len(t) > best_len:
            best, best_len = r.find(qn("w:rPr")), len(t)
    return (best if best is not None else first_rpr), first_rpr


def _has_numpr(p_el, doc_styles) -> bool:
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    sid = ppr.find(qn("w:pStyle")).get(qn("w:val")) if ppr is not None and ppr.find(qn("w:pStyle")) is not None else None
    seen = 0
    while sid and seen < 10:
        st = doc_styles.get(sid)
        if st is None:
            break
        sp = st.find(qn("w:pPr"))
        if sp is not None and sp.find(qn("w:numPr")) is not None:
            num = sp.find(qn("w:numPr")).find(qn("w:numId"))
            return num is None or num.get(qn("w:val")) != "0"
        b = st.find(qn("w:basedOn"))
        sid = b.get(qn("w:val")) if b is not None else None
        seen += 1
    return False


def _caps_in_style(p_el, rpr, doc_styles) -> bool:
    for el in [rpr]:
        if el is not None and (el.find(qn("w:caps")) is not None or el.find(qn("w:smallCaps")) is not None):
            return True
    ppr = p_el.find(qn("w:pPr"))
    sid = ppr.find(qn("w:pStyle")).get(qn("w:val")) if ppr is not None and ppr.find(qn("w:pStyle")) is not None else None
    seen = 0
    while sid and seen < 10:
        st = doc_styles.get(sid)
        if st is None:
            break
        r = st.find(qn("w:rPr"))
        if r is not None and (r.find(qn("w:caps")) is not None or r.find(qn("w:smallCaps")) is not None):
            return True
        b = st.find(qn("w:basedOn"))
        sid = b.get(qn("w:val")) if b is not None else None
        seen += 1
    return False


def text_case(t: str) -> str:
    letters = [c for c in t if c.isalpha()]
    if letters and all(c.isupper() for c in letters):
        return "upper"
    words = [w for w in re.findall(r"[A-Za-z][\w'-]*", t)]
    big = [w for w in words if len(w) > 3]
    if big and sum(w[0].isupper() for w in big) / len(big) > 0.8 and len(big) > 1:
        return "title"
    if words and words[0][0].isupper():
        return "sentence"
    return "as-is"


def heading_scheme(texts: Dict[int, List[str]]):
    """From original heading texts infer scheme + trailing dots."""
    out = {}
    l1 = texts.get(1, [])
    l2 = texts.get(2, [])
    if any(re.match(r"^\s*[IVX]+\.\s", t) for t in l1):
        out["scheme"] = "ieee"
    elif any(re.match(r"^\s*\d+\.?\s", t) for t in l1):
        out["scheme"] = "decimal"
        out["decimal_trailing_dot"] = any(re.match(r"^\s*\d+\.\s", t) for t in l1)
        if l2:
            out["subsection_trailing_dot"] = any(re.match(r"^\s*\d+\.\d+\.\s", t) for t in l2)
    elif l1:
        out["scheme"] = "none"
    return out


def detect_ref_style(texts: List[str]) -> Optional[str]:
    texts = [t for t in texts if len(t) > 25][:15]
    if not texts:
        return None
    score = collections.Counter()
    for t in texts:
        s = t.strip()
        bracket = bool(re.match(r"^\[\d+\]", s))
        numdot = bool(re.match(r"^\d+\.\s", s))
        body = re.sub(r"^\s*(?:\[\d+\]|\d+\.)\s*", "", s)
        if bracket and re.search(r"[“\"].+?,[”\"]", body):
            score["ieee"] += 2
        elif bracket:
            score["elsevier-numeric"] += 1
            if re.search(r"\(\d{4}\)\s*\d", body):
                score["elsevier-numeric"] += 1
        if re.match(r"^[^:]{3,200}:\s", body) and re.search(r"\((?:19|20)\d{2}\)\.?\s*(?:https?://\S+)?$", body):
            score["lncs"] += 2
        if re.search(r"\b(?:19|20)\d{2};\s*\d", body):
            score["vancouver"] += 2
        if re.search(r"^[^;]+,\s*[A-Z]\.(?:[A-Z]\.)*;\s", body):
            score["mdpi"] += 2
        if re.search(r"\((?:19|20)\d{2}[a-z]?\)\.\s", body) and re.search(r",\s*&\s|, [A-Z]\. \(", body):
            score["apa"] += 2
        elif re.search(r"\((?:19|20)\d{2}[a-z]?\)\.\s", body):
            score["apa"] += 1
        if re.search(r"\((?:19|20)\d{2}[a-z]?\)\s*[‘']", body):
            score["harvard"] += 2
        if re.search(r"\((?:19|20)\d{2}[a-z]?\),\s*[“\"]", body) or re.search(r"\bVol\.\s*\d+\s+No\.", body):
            score["emerald-harvard"] += 2
        if re.match(r"^[A-Z][\w'-]+\s[A-Z]{1,3}(?:,|\s\()", body) and re.search(r"\((?:19|20)\d{2}\)\s[A-Z]", body):
            score["springer-basic" if not (bracket or numdot) else "springer-vancouver"] += 2
        if re.search(r"\.\s(?:19|20)\d{2}\.\s[“\"]", body):
            score["chicago"] += 2
    if not score:
        return None
    return score.most_common(1)[0][0]


def detect_citation_mode_text(text: str) -> Optional[str]:
    nb = len(BRACKET_RE.findall(text))
    na = len(PAREN_RE.findall(text)) + len(NARR_RE.findall(text))
    if nb == na == 0:
        return None
    if nb >= na:
        return "numeric"
    if re.search(r"\([A-Z][\w-]+(?: et al\.)? (?:19|20)\d{2}[a-z]?[;)]", text) and not re.search(
            r"\([A-Z][\w-]+(?: et al\.)?, (?:19|20)\d{2}", text):
        return "author-year-nocomma"
    return "author-year"


def _label_info(text: str, regex) -> Optional[dict]:
    m = regex.match(text)
    if not m:
        return None
    rest = m.group("rest")
    lab = text[: len(text) - len(rest)] if rest else text
    lab_word = re.match(r"\s*([A-Za-z][A-Za-z ]*[A-Za-z])", lab).group(1)
    suffix = lab[len(lab.split(lab_word, 1)[0]) + len(lab_word):]
    return {"label": lab_word, "suffix": suffix if suffix.strip() else (suffix or " "), "rest": rest}


def caption_info(texts: List[str], kind: str) -> dict:
    out = {}
    for t in texts:
        m = re.match(r"^\s*(Fig\.?|Figure|FIGURE|FIG\.?|Table|TABLE|Tab\.?)\s*(\d+|[IVXLC]+)(\s*[.:—–\-|]?\s*)(.*)$", t)
        if not m:
            continue
        label, num, sep, rest = m.groups()
        if kind == "figure":
            out["figure_label"] = label[0] + label[1:].lower() if label.isupper() else label
        else:
            out["table_label"] = label.capitalize() if label.isupper() else label
            out["table_label_case"] = "upper" if label.isupper() else "as-is"
            out["table_numbering"] = "roman" if not num.isdigit() else "arabic"
            out["table_label_newline"] = not rest.strip()
        sep = re.sub(r"\s+", " ", sep)
        out["separator"] = sep if sep.strip() else " "
        if sep.strip() and not sep.endswith(" "):
            out["separator"] = sep + " "
        break
    return out


# ------------------------------------------------------------------ DOCX template
def analyze_docx_template(data: bytes):
    """Return (spec_overrides, TemplateProfile)."""
    import docx
    from .read_docx import DocxReader

    rdr = DocxReader(io.BytesIO(data))
    doc = rdr.read()
    d = rdr.d
    styles = {s.get(qn("w:styleId")): s for s in d.styles.element.findall(qn("w:style"))}
    prof = TemplateProfile(data=data)
    ov: dict = {"page": {}, "font": {}, "title": {}, "abstract": {}, "keywords": {}, "headings": {},
                "captions": {}, "references": {}, "citations": {}}

    # ---- sections
    body = d.element.body
    sect_list = []
    for p in body.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            sect_list.append(ppr.find(qn("w:sectPr")))
    final = body.find(qn("w:sectPr"))
    if final is not None:
        sect_list.append(final)
    prof.sections = [_x(s) for s in sect_list]

    def sect_info(s):
        info = {}
        pg = s.find(qn("w:pgSz")); mar = s.find(qn("w:pgMar")); cols = s.find(qn("w:cols"))
        if pg is not None:
            w, h = int(pg.get(qn("w:w"))), int(pg.get(qn("w:h")))
            info["size"] = "A4" if abs(w - 11906) < 200 else ("Letter" if abs(w - 12240) < 200 else f"{w}x{h}")
        if mar is not None:
            g = lambda k: round(int(mar.get(qn(k), "1440")) * TWIP_CM, 2)
            info["margins_cm"] = [g("w:top"), g("w:bottom"), g("w:left"), g("w:right")]
        n = int(cols.get(qn("w:num"), "1")) if cols is not None else 1
        info["columns"] = n
        if cols is not None and cols.get(qn("w:space")):
            info["col_gap_cm"] = round(int(cols.get(qn("w:space"))) * TWIP_CM, 2)
        return info

    if sect_list:
        ov["page"].update(sect_info(sect_list[-1]))
        if len(sect_list) > 1:
            first = sect_info(sect_list[0])
            ov["page"]["front_matter_single_column"] = first.get("columns", 1) == 1 and ov["page"]["columns"] > 1
            prof.front_section = _x(sect_list[0])
        prof.found.append(f"page: {ov['page'].get('size')} · {ov['page'].get('columns')} column(s) · "
                          f"margins {ov['page'].get('margins_cm')} cm")

    # ---- exemplars
    heading_texts = collections.defaultdict(list)
    ref_texts, para_texts = [], []
    first_para_after_heading = None
    prev = None
    cap_seen = {"figure": [], "table": []}
    table_after_caption = None
    for b in doc.blocks:
        el = b.src
        if b.role == M.TABLE and el is not None:
            tp = el.find(qn("w:tblPr"))
            if prof.table_tblPr is None and tp is not None:
                prof.table_tblPr = _x(tp)
            if table_after_caption is None:
                table_after_caption = prev is not None and prev.role == M.CAPTION and prev.cap_kind == "table"
            prev = b
            continue
        if el is None or etree.QName(el).localname != "p":
            prev = b
            continue
        key = None
        if b.role == M.HEADING:
            key = f"h{min(b.level, 4)}"
            heading_texts[b.level].append(b.orig)
        elif b.role == M.CAPTION:
            key = f"caption_{b.cap_kind}"
            cap_seen[b.cap_kind].append(b.orig)
        elif b.role == M.LIST:
            key = "list_number" if b.ordered else "list_bullet"
        elif b.role == M.PARA:
            para_texts.append(b.orig)
            if len(b.orig) > 120 and "para" not in prof.exemplars:
                key = "para"
            elif "para_short" not in prof.exemplars:
                key = "para_short"
            if prev is not None and prev.role == M.HEADING and first_para_after_heading is None:
                first_para_after_heading = b
                prof.exemplars["para_first"] = None  # placeholder filled below
                key2 = "para_first"
                dom, first = _dominant_rpr(el)
                prof.exemplars[key2] = {"pPr": _x(el.find(qn("w:pPr"))), "rPr": _x(dom), "label_rPr": None,
                                        "auto_num": False}
        elif b.role == M.REFERENCE:
            ref_texts.append(b.orig)
            key = "reference"
        elif b.role in (M.TITLE, M.AUTHOR, M.AFFIL, M.ABSTRACT, M.KEYWORDS, M.EQUATION, M.FIGURE,
                        M.ABSTRACT_HEADING, M.OTHER_FRONT):
            key = b.role
        if key and key not in prof.exemplars:
            labelled = b.role in (M.ABSTRACT, M.KEYWORDS, M.CAPTION)
            dom, first = _dominant_rpr(el, skip_label=labelled)
            prof.exemplars[key] = {
                "pPr": _x(el.find(qn("w:pPr"))),
                "rPr": _x(dom),
                "label_rPr": _x(first) if labelled else None,
                "auto_num": _has_numpr(el, styles),
                "caps": _caps_in_style(el, dom, styles),
                "text": b.orig,
            }
        prev = b

    # ---- derive spec overrides
    pe = prof.exemplars.get("para") or prof.exemplars.get("para_short")
    if pe and pe.get("rPr"):
        r = etree.fromstring(pe["rPr"])
        fonts = r.find(qn("w:rFonts"))
        if fonts is not None and fonts.get(qn("w:ascii")):
            ov["font"]["family"] = fonts.get(qn("w:ascii"))
        sz = r.find(qn("w:sz"))
        if sz is not None:
            ov["font"]["size"] = int(sz.get(qn("w:val"))) / 2
    if "family" not in ov["font"]:
        try:
            nf = d.styles["Normal"].font
            if nf.name:
                ov["font"]["family"] = nf.name
        except KeyError:
            pass
    if "size" not in ov["font"]:
        ov["font"]["size"] = rdr.default_size
    prof.found.append(f"body font: {ov['font'].get('family', '?')} {ov['font'].get('size')} pt")

    # headings
    hs = heading_scheme(heading_texts)
    for lvl in (1, 2, 3):
        ex = prof.exemplars.get(f"h{lvl}")
        if ex and ex.get("auto_num"):
            hs.setdefault("auto_levels", []).append(lvl)
    if hs:
        ov["headings"].update(hs)
        levels = []
        for lvl in (1, 2, 3, 4):
            ts = heading_texts.get(lvl)
            ex = prof.exemplars.get(f"h{lvl}", {})
            if ts:
                bare = strip_heading_number(ts[0])[1]
                case = "as-is" if ex.get("caps") else text_case(bare)
                levels.append({"case": case if case != "sentence" or len(bare.split()) > 1 else "as-is"})
            else:
                levels.append({})
        ov["headings"]["levels"] = levels
        prof.found.append(f"headings: {hs.get('scheme')} numbering" +
                          (" (Word auto-numbered)" if hs.get("auto_levels") else ""))

    # abstract / keywords labels
    ab = prof.exemplars.get(M.ABSTRACT)
    if M.ABSTRACT_HEADING in prof.exemplars:
        ov["abstract"].update(layout="block", heading=re.sub(r"[\s:—–.-]+$", "", prof.exemplars[M.ABSTRACT_HEADING]["text"]) or "Abstract")
        prof.found.append("abstract: separate heading line")
    elif ab and ab.get("text"):
        li = _label_info(ab["text"], ABSTRACT_RE)
        if li and li["rest"]:
            ov["abstract"].update(layout="inline", heading=li["label"], suffix=li["suffix"])
            prof.found.append(f"abstract: inline label “{li['label']}{li['suffix'].strip() or ' '}”")
    kw = prof.exemplars.get(M.KEYWORDS)
    if kw and kw.get("text"):
        li = _label_info(kw["text"], KEYWORDS_RE)
        if li:
            rest = li["rest"]
            sep = "; " if rest.count(";") >= max(1, rest.count(",")) else (" · " if "·" in rest else ", ")
            if "—" in li["suffix"]:
                suffix = "—"
            else:
                suffix = li["suffix"]
            ov["keywords"].update(label=li["label"].strip(), suffix=suffix, separator=sep,
                                  end_period=rest.strip().endswith("."))
            prof.found.append(f"keywords label “{li['label']}”, separator “{sep.strip()}”")

    # captions
    ci = caption_info(cap_seen["figure"], "figure")
    ci_t = caption_info(cap_seen["table"], "table")
    sep = ci.get("separator") or ci_t.get("separator")
    ov["captions"].update({k: v for k, v in {**ci_t, **ci}.items() if k != "separator"})
    if sep:
        ov["captions"]["separator"] = sep
    if table_after_caption is not None:
        ov["captions"]["table_position"] = "above" if table_after_caption else "below"
    if ci or ci_t:
        prof.found.append(f"captions: “{ov['captions'].get('figure_label', 'Figure')} 1{ov['captions'].get('separator', '. ').rstrip()}”"
                          f", tables “{ov['captions'].get('table_label', 'Table')} …” {ov['captions'].get('table_position', 'above')}")

    # references & citations
    rs = detect_ref_style(ref_texts)
    if rs:
        ov["references"]["style"] = rs
        prof.found.append(f"reference style: {rs}")
    cm = detect_citation_mode_text(" ".join(para_texts))
    if cm:
        ov["citations"]["mode"] = cm
    elif rs:
        from .refs import NUMERIC_STYLES
        ov["citations"]["mode"] = "numeric" if rs in NUMERIC_STYLES else "author-year"
    for b in doc.blocks:
        if b.role == M.HEADING and re.sub(r"[^a-z]", "", b.text.lower()) in ("references", "bibliography", "literaturecited"):
            ov["references"]["heading"] = strip_heading_number(b.orig)[1]
            break
    # title case
    te = prof.exemplars.get(M.TITLE)
    ov = {k: v for k, v in ov.items() if v}
    return ov, prof


# ------------------------------------------------------------------ PDF sample paper
PDF_FONT_MAP = [
    (r"times|nimbusrom|tex-?gyre-?termes|stix|liberationserif", "Times New Roman"),
    (r"palatino|pagella|urwpalladio", "Palatino Linotype"),
    (r"arial|helvetica|nimbussan|liberationsans", "Arial"),
    (r"calibri", "Calibri"), (r"cambria", "Cambria"), (r"georgia", "Georgia"),
    (r"cmr|lmroman|computermodern|sfrm", "Computer Modern"), (r"garamond", "Garamond"),
    (r"minion", "Minion Pro"), (r"charter", "Charter"),
]


def map_pdf_font(name: str) -> str:
    n = re.sub(r"^[A-Z]{6}\+", "", name or "").lower().replace(" ", "")
    for pat, fam in PDF_FONT_MAP:
        if re.search(pat, n):
            return fam
    return re.sub(r"[-,].*$", "", re.sub(r"^[A-Z]{6}\+", "", name or "")) or "Times New Roman"


def analyze_pdf_sample(data: bytes) -> dict:
    """Infer spec overrides from a published sample paper PDF."""
    import pymupdf
    from .read_pdf import read_pdf

    pdf = pymupdf.open(stream=data, filetype="pdf")
    ov: dict = {"page": {}, "font": {}, "title": {}, "headings": {}, "abstract": {}, "keywords": {},
                "captions": {}, "references": {}, "citations": {}, "_found": []}
    p0 = pdf[min(1, len(pdf) - 1)]
    w, h = p0.rect.width, p0.rect.height
    ov["page"]["size"] = "A4" if abs(w - 595) < 8 else ("Letter" if abs(w - 612) < 8 else f"{w:.0f}x{h:.0f}pt")
    xs0, xs1, ys0, ys1 = [], [], [], []
    bottoms = []
    left_n = right_n = full_n = 0
    for pg in list(pdf)[1: min(len(pdf), 5)] or list(pdf)[:1]:
        for b in pg.get_text("blocks"):
            x0, y0, x1, y1, t = b[:5]
            if len(t.strip()) < 40:
                continue
            xs0.append(x0); xs1.append(x1); ys0.append(y0); ys1.append(y1)
            if x1 < pg.rect.width / 2 + 10:
                left_n += 1
            elif x0 > pg.rect.width / 2 - 10:
                right_n += 1
            else:
                full_n += 1
    if len(pdf) > 1:
        for pg in list(pdf)[:-1]:
            ys = [b[3] for b in pg.get_text("blocks") if len(b[4].strip()) > 40 and b[3] < pg.rect.height * 0.93]
            if ys:
                bottoms.append(max(ys))
    if xs0:
        pt_cm = 2.54 / 72
        bottom = max(bottoms) if bottoms else max(ys1)
        ov["page"]["margins_cm"] = [round(min(ys0) * pt_cm, 2), round((h - bottom) * pt_cm, 2),
                                    round(min(xs0) * pt_cm, 2), round((w - max(xs1)) * pt_cm, 2)]
    ov["page"]["columns"] = 2 if left_n >= 3 and right_n >= 3 and (left_n + right_n) > full_n else 1
    # fonts
    fonts = collections.Counter(); sizes = collections.Counter()
    for pg in pdf:
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    fonts[s["font"]] += len(s["text"]); sizes[round(s["size"] * 2) / 2] += len(s["text"])
    if fonts:
        ov["font"]["family"] = map_pdf_font(fonts.most_common(1)[0][0])
        ov["font"]["size"] = sizes.most_common(1)[0][0]
    ov["_found"].append(f"page: {ov['page']['size']} · {ov['page']['columns']} column(s)")
    ov["_found"].append(f"body font: {ov['font'].get('family')} {ov['font'].get('size')} pt")

    doc = read_pdf(data)
    t = doc.first(M.TITLE)
    if t:
        ov["title"]["size"] = round(t.src_size)
        ov["title"]["bold"] = t.src_bold
    htexts = collections.defaultdict(list)
    hsizes = collections.defaultdict(list)
    for b in doc.blocks:
        if b.role == M.HEADING:
            htexts[b.level].append(b.orig)
            hsizes[b.level].append((b.src_size, b.src_bold))
    hs = heading_scheme(htexts)
    if hs:
        ov["headings"].update(hs)
        levels = []
        for lvl in (1, 2, 3, 4):
            if htexts.get(lvl):
                bare = strip_heading_number(htexts[lvl][0])[1]
                sz, bold = hsizes[lvl][0]
                levels.append({"case": text_case(bare) if text_case(bare) == "upper" else "as-is",
                               "size": round(sz * 2) / 2, "bold": bold})
            else:
                levels.append({})
        ov["headings"]["levels"] = levels
        ov["_found"].append(f"headings: {hs.get('scheme')} numbering")
    if doc.first(M.ABSTRACT_HEADING):
        ov["abstract"]["layout"] = "block"
    else:
        ab = doc.first(M.ABSTRACT)
        if ab:
            li = _label_info(ab.orig, ABSTRACT_RE)
            if li and li["rest"]:
                ov["abstract"].update(layout="inline", heading=li["label"], suffix=li["suffix"])
    kw = doc.first(M.KEYWORDS)
    if kw:
        li = _label_info(kw.orig, KEYWORDS_RE)
        if li:
            rest = li["rest"]
            sep = "; " if rest.count(";") >= max(1, rest.count(",")) else (" · " if "·" in rest else ", ")
            ov["keywords"].update(label=li["label"].strip(), suffix="—" if "—" in li["suffix"] else li["suffix"],
                                  separator=sep)
    figs = [b.orig for b in doc.blocks if b.role == M.CAPTION and b.cap_kind == "figure"]
    tabs = [b.orig for b in doc.blocks if b.role == M.CAPTION and b.cap_kind == "table"]
    ci = {**caption_info(tabs, "table"), **caption_info(figs, "figure")}
    ov["captions"].update(ci)
    rs = detect_ref_style([b.orig for b in doc.blocks if b.role == M.REFERENCE])
    if rs:
        ov["references"]["style"] = rs
        ov["_found"].append(f"reference style: {rs}")
    cm = detect_citation_mode_text(" ".join(b.orig for b in doc.blocks if b.role == M.PARA))
    if cm:
        ov["citations"]["mode"] = cm
    found = ov.pop("_found")
    ov = {k: v for k, v in ov.items() if v}
    ov["_found"] = found
    return ov
