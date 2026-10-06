"""Reference parsing, reference-list styling and in-text citation conversion."""
from __future__ import annotations

import copy
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from . import model as M
from .model import Run

DOI_RE = re.compile(r"(?:https?://(?:dx\.)?doi\.org/|doi:\s*|DOI:\s*)?(10\.\d{4,9}/[^\s,;]+)", re.I)
URL_RE = re.compile(r"https?://\S+")
YEAR_RE = re.compile(r"\b((?:19|20)\d{2})([a-z])?\b")
LABEL_RE = re.compile(r"^\s*(?:\[\d{1,3}\]|\(\d{1,3}\)|\d{1,3}\.(?=\s)|\d{1,3}\)(?=\s))\s*")
NAME = r"[A-ZÀ-ſ][\wÀ-ſ'’\-]+"
PARTICLE = r"(?:(?:van|von|der|de|da|del|di|la|le|du|den|ter|bin|al|el)\s+)*"


@dataclass
class Ref:
    raw: str
    authors: List[Tuple[str, str]] = field(default_factory=list)  # (family, initials "A.B.")
    etal: bool = False
    year: str = ""
    suffix: str = ""
    title: str = ""
    container: str = ""
    volume: str = ""
    issue: str = ""
    pages: str = ""
    doi: str = ""
    url: str = ""
    publisher: str = ""
    kind: str = "journal"
    confident: bool = True
    old_num: Optional[int] = None

    @property
    def first_family(self) -> str:
        return self.authors[0][0] if self.authors else ""


def _fold(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).lower()


# ------------------------------------------------------------------ author parsing
def _initials(given: str) -> str:
    parts = re.split(r"[\s.]+", given.strip())
    out = []
    for p in parts:
        if not p:
            continue
        if "-" in p:
            out.append("-".join(x[:1].upper() + "." for x in p.split("-") if x))
        elif p.isupper() and len(p) <= 3 and len(parts) == 1:   # Vancouver "AB"
            out.extend(ch + "." for ch in p)
        else:
            out.append(p[0].upper() + ".")
    return "".join(out)


def parse_authors(s: str) -> Tuple[List[Tuple[str, str]], bool]:
    s = s.strip().rstrip(".,:;")
    etal = bool(re.search(r"\bet\s+al\.?", s))
    s = re.sub(r",?\s*\bet\s+al\.?", "", s)
    s = s.replace("&", ",").replace(" and ", ", ").replace(";", ",")
    s = re.sub(r"\s+", " ", s).strip(" ,")
    if not s:
        return [], etal
    if re.search(r"\b(?:organi[sz]ation|association|institute|society|council|ministry|agency|committee|"
                 r"commission|bureau|office|group|consortium|foundation|nations|union|board|authority|"
                 r"corporation|inc|ltd|university)\b", s, re.I) and not re.search(r",\s*[A-Z]\.", s):
        return [(s, "")], etal
    cands = []
    # APA/Harvard: Smith, A. B.
    apa = re.findall(r"(" + PARTICLE + NAME + r"(?:\s" + NAME + r")?),\s*((?:[A-Z][a-z]?\.?\s?-?\s?){1,4})(?=,|$)", s)
    cands.append([(f, _initials(g)) for f, g in apa])
    # IEEE/Elsevier: A. B. Smith
    ieee = re.findall(r"((?:[A-Z][a-z]?\.\s?-?){1,4})\s*(" + PARTICLE + NAME + r"(?:\s" + NAME + r")?)(?=,|$)", s)
    cands.append([(f, _initials(g)) for g, f in ieee])
    # Vancouver/Springer: Smith AB
    van = re.findall(r"(" + PARTICLE + NAME + r"(?:\s" + NAME + r")?)\s([A-Z]{1,4})(?=,|$)", s)
    cands.append([(f, _initials(g)) for f, g in van])
    # Full names: Ravi Kumar
    parts = [p.strip() for p in s.split(",") if p.strip()]
    full = []
    for p in parts:
        w = p.split()
        if 2 <= len(w) <= 4 and all(re.match(r"[A-ZÀ-ſ]", x) for x in w):
            full.append((w[-1], _initials(" ".join(w[:-1]))))
    cands.append(full)
    # Chicago / MLA: "Family, Given, Given Family, and Given Family"
    chi = []
    m = re.match(r"^(" + PARTICLE + NAME + r"),\s*((?:" + NAME + r"|[A-Z]\.)(?:\s(?:" + NAME + r"|[A-Z]\.))*)(?:,|$)", s)
    if m:
        chi.append((m.group(1), _initials(m.group(2))))
        for part in [x.strip() for x in s[m.end():].split(",") if x.strip()]:
            w = part.split()
            if 2 <= len(w) <= 4 and all(re.match(r"[A-Z\u00C0-\u017F]", x) for x in w):
                chi.append((w[-1], _initials(" ".join(w[:-1]))))
            else:
                chi = []
                break
    cands.append(chi)
    best = max(cands, key=len)
    # Single "Surname, Given Names" (APA-style with full given name)
    if not best:
        m = re.match(r"^(" + NAME + r"),\s*(" + NAME + r"(?:\s" + NAME + r")*)$", s)
        if m:
            best = [(m.group(1), _initials(m.group(2)))]
        elif re.fullmatch(NAME + r"(?:\s" + NAME + r")*", s):
            best = [(s, "")]    # organisation as author
    return best, etal


# ------------------------------------------------------------------ container parsing
def _parse_container(rest: str, ref: Ref):
    r = rest.strip()
    r = re.sub(r"\s*(?:Retrieved from|Available (?:at|online)).*$", "", r, flags=re.I)
    m = re.search(r"\bvol(?:ume)?\.?\s*(\d+)", r, re.I)
    if m:
        ref.volume = m.group(1)
    m2 = re.search(r"\b(?:no|issue|iss)\.?\s*(\d+[A-Za-z]?(?:[-–/]\d+)?)", r, re.I)
    if m2:
        ref.issue = m2.group(1)
    m3 = re.search(r"\bpp?\.?\s*(e?\d+\s*[-–—]\s*e?\d+|e?\d+)", r)
    if m3:
        ref.pages = m3.group(1).replace(" ", "")
    if not ref.pages:
        mp = re.search(r"[,:]\s*(e?\d+\s*[-–—]\s*e?\d+)\s*\.?\s*$", r)
        if mp:
            ref.pages = mp.group(1).replace(" ", "")
    if not ref.volume:
        m = re.search(r"(?<![\d.])(\d{1,4})\s*\((\d+[A-Za-z]?(?:[-–/]\d+)?)\)\s*[:,]?\s*(e?\d+(?:\s*[-–—]\s*e?\d+)?)?", r)
        if m:
            ref.volume, ref.issue = m.group(1), m.group(2)
            if m.group(3) and not ref.pages:
                ref.pages = m.group(3).replace(" ", "")
        else:
            m = re.search(r"(?<![\d.])(\d{1,4})\s*:\s*(e?\d+(?:\s*[-–—]\s*e?\d+)?)", r)
            if m:
                ref.volume, ref.pages = m.group(1), ref.pages or m.group(2).replace(" ", "")
            else:
                m = re.search(r",\s*(\d{1,4})\s*,\s*(e?\d+(?:\s*[-–—]\s*e?\d+)?)\s*\.?\s*$", r)
                if m:
                    ref.volume, ref.pages = m.group(1), ref.pages or m.group(2).replace(" ", "")
    # container name: text up to first volume/number marker
    cut = re.search(r",?\s*(?:\bvol\.|\bvolume\b|\bno\.|\bpp?\.|(?<![\w.])\d{1,4}\s*[(:,]|\d{4}\s*[;,]|(?<![\w.])\d{1,4}\s*$|(?<![\w.])e?\d+\s*[-–—]\s*e?\d+)", r, re.I)
    name = r[: cut.start()] if cut else r
    name = re.sub(r"^\s*(?:In:?|in)\s+", "", name).strip(" .,;:")
    name = YEAR_RE.sub("", name).strip(" .,;:()")
    ref.container = name
    if re.search(r"proceedings|conference|symposium|workshop|\bconf\.|\bproc\.", name, re.I):
        ref.kind = "conference"
    if ref.pages:
        ref.pages = re.sub(r"\s*[-–—]\s*", "–", ref.pages)


def _split_title(rest: str) -> Tuple[str, str]:
    """Title is the first sentence of `rest` (keep ?/! endings), or the quoted string."""
    q = re.match(r"^\s*[\"“‘](.+?)[,.]?[\"”’]\s*[,.:]?\s*(.*)$", rest)
    if q:
        return q.group(1).strip(), q.group(2)
    m = re.match(r"^\s*[\"“'‘]?(.+?)[\"”'’]?([.?!])[\"”'’]?\s+(.*)$", rest)
    if not m:
        return rest.strip(" ."), ""
    title = m.group(1)
    if m.group(2) in "?!":
        title += m.group(2)
    return title.strip(), m.group(3)


def _author_like(a: str) -> bool:
    """True when `a` is (almost) entirely author names."""
    if re.search(r"\b(?:19|20)\d{2}\b|[“”\"]", a):
        return False
    auths, _ = parse_authors(a)
    if not auths:
        return False
    fam_words = {w.lower() for f, _ in auths for w in f.split()}
    words = re.findall(r"[A-Za-z\u00C0-\u017F'’\-]+", a)
    unexplained = [w for w in words if len(w.rstrip(".")) > 2 and w.lower() not in fam_words
                   and w.lower() not in ("and", "et", "al", "eds", "ed")]
    # given names written out in full are explained by the parsed initials
    ini = {ch for _, i in auths for ch in i if ch.isalpha()}
    unexplained = [w for w in unexplained if not (w[0].isupper() and w[0] in ini)]
    return len(unexplained) <= max(1, len(words) // 5)


def parse_reference(raw: str) -> Ref:
    text = re.sub(r"\s+", " ", raw).strip()
    ref = Ref(raw=text)
    lm = LABEL_RE.match(text)
    if lm:
        n = re.search(r"\d+", lm.group(0))
        ref.old_num = int(n.group(0)) if n else None
        text = text[lm.end():]
    dm = DOI_RE.search(text)
    if dm:
        ref.doi = dm.group(1).rstrip(".")
        text = (text[: dm.start()] + text[dm.end():]).strip()
    um = URL_RE.search(text)
    if um:
        ref.url = um.group(0).rstrip(".")
        text = (text[: um.start()] + text[um.end():]).strip()
    text = re.sub(r"\s*(?:\[?Online\]?\.?|Accessed:?.*$|\bdoi:?\s*$)", " ", text, flags=re.I).strip(" .,")

    # IEEE: Authors, “Title,” rest
    m = re.match(r"^(?P<a>.+?),?\s*[\"“](?P<t>.+?)[,.]?[\"”],?\s*(?P<r>.*)$", text)
    if m and len(m.group("a")) < 400 and _author_like(m.group("a")):
        ref.authors, ref.etal = parse_authors(m.group("a"))
        ref.title = m.group("t").strip(" ,.")
        rest = re.sub(r"^\s*in\s+", "", m.group("r"))
        y = YEAR_RE.findall(rest)
        if y:
            ref.year, ref.suffix = y[-1]
        _parse_container(rest, ref)
        return _finish(ref)
    # Chicago author-date: Authors. 2017. “Title.” Journal 19 (3): 1–30.
    m = re.match(r"^(?P<a>.+?)\.\s+(?P<y>(?:19|20)\d{2})(?P<s>[a-z])?\.\s+[\"“](?P<t>.+?)[.,]?[\"”]\s*(?P<r>.*)$", text)
    if m and _author_like(m.group("a")):
        ref.authors, ref.etal = parse_authors(m.group("a"))
        ref.year, ref.suffix = m.group("y"), m.group("s") or ""
        ref.title = m.group("t").strip(" .,")
        rest = re.sub(r"^\s*In\s+", "", m.group("r"))
        _parse_container(rest, ref)
        return _finish(ref)
    # LNCS: Authors: Title. Journal 9(2), 11–20 (2020)
    m = re.match(r"^(?P<a>[^:]{3,300}?):\s+(?P<r>.+)$", text)
    if m and _author_like(m.group("a")):
        ref.authors, ref.etal = parse_authors(m.group("a"))
        ref.title, rest = _split_title(m.group("r"))
        y = YEAR_RE.findall(rest)
        if y:
            ref.year, ref.suffix = y[-1]
        _parse_container(rest, ref)
        return _finish(ref)
    # APA/Harvard: Authors (2020). Title. rest  |  Springer: Authors (2020) Title. rest
    m = re.match(r"^(?P<a>.+?)\s*\((?P<y>(?:19|20)\d{2})(?P<s>[a-z])?[^)]*\)[.,]?\s*(?P<r>.*)$", text)
    if m and len(m.group("a")) < 500 and _author_like(m.group("a")):
        ref.authors, ref.etal = parse_authors(m.group("a"))
        ref.year, ref.suffix = m.group("y"), m.group("s") or ""
        ref.title, rest = _split_title(m.group("r"))
        _parse_container(rest, ref)
        return _finish(ref)
    # Harvard without parentheses: Smith, J. 2020. Title.
    m = re.match(r"^(?P<a>.+?)[,.]?\s+(?P<y>(?:19|20)\d{2})(?P<s>[a-z])?[.,]\s+(?P<r>.*)$", text)
    if m and len(m.group("a")) < 300 and _author_like(m.group("a")):
        ref.authors, ref.etal = parse_authors(m.group("a"))
        ref.year, ref.suffix = m.group("y"), m.group("s") or ""
        ref.title, rest = _split_title(m.group("r"))
        _parse_container(rest, ref)
        return _finish(ref)
    # Vancouver / MDPI: Authors. Title. Journal. 2020;9(2):11-20.   |  Smith, A.; Lee, B. Title. J. 2020, 9, 1.
    m = re.match(r"^(?P<a>(?:[^.]|(?<=\b[A-Z])\.)+?)\.\s+(?P<r>.+)$", text)
    if m:
        a = m.group("a")
        auths, etal = parse_authors(a)
        if auths and _author_like(a):
            ref.authors, ref.etal = auths, etal
            ref.title, rest = _split_title(m.group("r"))
            y = YEAR_RE.findall(rest)
            if y:
                ref.year, ref.suffix = y[0]
            _parse_container(rest, ref)
            return _finish(ref)
    ref.confident = False
    y = YEAR_RE.findall(text)
    if y:
        ref.year, ref.suffix = y[0]
    ref.title = text
    return ref


def _finish(ref: Ref) -> Ref:
    if not ref.authors or not ref.title or not ref.year:
        ref.confident = False
    if ref.kind == "journal" and not ref.container and not ref.volume:
        ref.kind = "other"
    return ref


def enrich_crossref(ref: Ref, timeout=6.0) -> bool:
    """Fill/verify metadata from Crossref by DOI (or by title search). Returns True when updated."""
    try:
        import requests
        hdr = {"User-Agent": "FormatMatch/1.0 (mailto:info@cssoftwaresolutions.tech)"}
        if ref.doi:
            r = requests.get(f"https://api.crossref.org/works/{ref.doi}", headers=hdr, timeout=timeout)
            if r.status_code != 200:
                return False
            it = r.json()["message"]
        else:
            q = ref.title if ref.title and len(ref.title) > 20 else ref.raw
            r = requests.get("https://api.crossref.org/works", params={"query.bibliographic": q[:300], "rows": 1},
                             headers=hdr, timeout=timeout)
            items = r.json()["message"]["items"]
            if not items:
                return False
            it = items[0]
            import difflib
            t = (it.get("title") or [""])[0]
            if difflib.SequenceMatcher(None, _fold(t), _fold(ref.title or ref.raw[:len(t) + 40])).ratio() < 0.75:
                return False
        auth = []
        for a in it.get("author", []):
            if a.get("family"):
                auth.append((a["family"], _initials(a.get("given", ""))))
            elif a.get("name"):
                auth.append((a["name"], ""))
        if auth:
            ref.authors, ref.etal = auth, False
        ref.title = (it.get("title") or [ref.title])[0]
        ref.container = (it.get("container-title") or [ref.container])[0]
        ref.volume = it.get("volume", ref.volume) or ref.volume
        ref.issue = it.get("issue", ref.issue) or ref.issue
        ref.pages = (it.get("page") or it.get("article-number") or ref.pages or "").replace("-", "–")
        dp = (it.get("published-print") or it.get("published-online") or it.get("issued") or {}).get("date-parts")
        if dp and dp[0] and dp[0][0]:
            ref.year = str(dp[0][0])
        ref.doi = it.get("DOI", ref.doi)
        t = it.get("type", "")
        ref.kind = "conference" if "proceedings" in t else ("book" if "book" in t else "journal")
        ref.confident = True
        return True
    except Exception:
        return False


# ------------------------------------------------------------------ name formatting helpers
def _ini(i: str, sep="") -> str:
    """'A.B.' -> 'A.B.' / 'A. B.' / 'AB'."""
    if sep == "none":
        return i.replace(".", "").replace("-", "")
    if sep == " ":
        return re.sub(r"\.(?=[A-Z])", ". ", i)
    return i


def _join(names: List[str], last_sep: str, sep=", ") -> str:
    if len(names) <= 1:
        return "".join(names)
    return sep.join(names[:-1]) + last_sep + names[-1]


def _authors(ref: Ref, style: str) -> str:
    A = ref.authors
    if not A:
        return ""
    if style == "ieee":
        names = [f"{_ini(i, ' ')} {f}".strip() for f, i in A]
        if len(names) > 6 or ref.etal:
            return names[0] + " et al."
        return _join(names, ", and " if len(names) > 2 else " and ")
    if style == "apa":
        names = [f"{f}, {_ini(i, ' ')}".strip(", ") for f, i in A]
        if len(names) > 20:
            names = names[:19] + ["…"] + names[-1:]
        s = _join(names, ", & " if len(names) > 2 else ", & ")
        return s + (", et al." if ref.etal else "")
    if style in ("harvard", "emerald-harvard"):
        names = [f"{f}, {_ini(i)}".strip(", ") for f, i in A]
        return _join(names, " and ") + (" et al." if ref.etal else "")
    if style in ("vancouver", "ama", "springer"):
        names = [f"{f} {_ini(i, 'none')}".strip() for f, i in A]
        limit = 6 if style == "vancouver" else (6 if style == "ama" else 99)
        if len(names) > limit or ref.etal:
            keep = 3 if style == "ama" else 6
            return ", ".join(names[:keep]) + ", et al"
        return ", ".join(names)
    if style == "chicago":
        first = f"{A[0][0]}, {_ini(A[0][1], ' ')}"
        rest = [f"{_ini(i, ' ')} {f}" for f, i in A[1:]]
        return _join([first] + rest, ", and " if len(A) > 2 else " and ")
    if style == "lncs":
        names = [f"{f}, {_ini(i)}" for f, i in A]
        return ", ".join(names) + (", et al." if ref.etal else "")
    if style == "mdpi":
        return "; ".join(f"{f}, {_ini(i)}" for f, i in A) + ("; et al." if ref.etal else "")
    if style == "elsevier":
        names = [f"{_ini(i)} {f}".strip() for f, i in A]
        return ", ".join(names) + (", et al." if ref.etal else "")
    if style == "mla":
        if len(A) == 1:
            return f"{A[0][0]}, {_ini(A[0][1], ' ')}"
        if len(A) == 2:
            return f"{A[0][0]}, {_ini(A[0][1], ' ')}, and {_ini(A[1][1], ' ')} {A[1][0]}"
        return f"{A[0][0]}, {_ini(A[0][1], ' ')}, et al"
    return ", ".join(f for f, _ in A)


def _end(s: str, ch=".") -> str:
    s = s.rstrip()
    return s if not s or s[-1] in ".?!" else s + ch


def format_reference(ref: Ref, style: str, n: int) -> List[Run]:
    """Return runs (with italics/bold) for one reference entry in `style`."""
    R = Run
    if not ref.confident and not ref.authors:
        body = LABEL_RE.sub("", ref.raw)
        return _with_label([R(body)], style, n)
    doi = f"https://doi.org/{ref.doi}" if ref.doi else (ref.url or "")
    t, c, v, i, p, y = ref.title, ref.container, ref.volume, ref.issue, ref.pages, ref.year + ref.suffix
    out: List[Run] = []
    if style == "ieee":
        out.append(R(f"{_authors(ref, 'ieee')}, “{t},” "))
        if c:
            out.append(R(("in " if ref.kind == "conference" else "")))
            out.append(R(c, italic=True))
        bits = []
        if v: bits.append(f"vol. {v}")
        if i: bits.append(f"no. {i}")
        if p: bits.append(("pp. " if "–" in p else "Art. no. " if p.startswith("e") or len(p) > 5 else "p. ") + p)
        if y: bits.append(y)
        out.append(R((", " if c else "") + ", ".join(bits)))
        if ref.doi:
            out.append(R(f", doi: {ref.doi}"))
        out.append(R("."))
    elif style == "apa":
        out.append(R(f"{_authors(ref, 'apa')} ({y or 'n.d.'}). {_end(t)} "))
        if c:
            out.append(R(c, italic=True))
            if v:
                out.append(R(", "))
                out.append(R(v, italic=True))
            if i:
                out.append(R(f"({i})"))
            if p:
                out.append(R(f", {p}"))
            out.append(R("."))
        if doi:
            out.append(R(f" {doi}"))
    elif style == "harvard":
        out.append(R(f"{_authors(ref, 'harvard')} ({y}) ‘{t}’, "))
        if c:
            out.append(R(c, italic=True))
        bits = []
        if v: bits.append(v + (f"({i})" if i else ""))
        if p: bits.append(f"pp. {p}" if "–" in p else f"p. {p}")
        out.append(R((", " if bits else "") + ", ".join(bits) + "."))
        if ref.doi:
            out.append(R(f" doi: {ref.doi}."))
    elif style == "emerald-harvard":
        out.append(R(f"{_authors(ref, 'emerald-harvard')} ({y}), “{t}”, "))
        if c:
            out.append(R(c, italic=True))
        bits = []
        if v: bits.append(f"Vol. {v}" + (f" No. {i}" if i else ""))
        if p: bits.append(f"pp. {p.replace('–', '-')}")
        if ref.doi: bits.append(f"doi: {ref.doi}")
        out.append(R((", " if bits else "") + ", ".join(bits) + "."))
    elif style in ("vancouver", "ama"):
        out.append(R(f"{_authors(ref, style)}. {_end(t)} "))
        if c:
            out.append(R(_end(c), italic=(style == "ama")))
        loc = f" {y}" + (f";{v}" if v else "") + (f"({i})" if i else "") + (f":{p}" if p else "") + "."
        out.append(R(loc))
        if ref.doi:
            out.append(R(f" doi:{ref.doi}"))
    elif style == "chicago":
        out.append(R(f"{_end(_authors(ref, 'chicago'))} {y}. “{_end(t)}” "))
        if c:
            out.append(R(c, italic=True))
        out.append(R((f" {v}" if v else "") + (f" ({i})" if i else "") + (f": {p}" if p else "") + "."))
        if doi:
            out.append(R(f" {doi}."))
    elif style in ("springer-basic", "springer-vancouver"):
        out.append(R(f"{_authors(ref, 'springer')} ({y}) {_end(t)} "))
        if c:
            out.append(R(c))
        out.append(R((f" {v}" if v else "") + (f"({i})" if i else "") + (f":{p}" if p else "")))
        if doi:
            out.append(R(f". {doi}"))
    elif style == "lncs":
        out.append(R(f"{_authors(ref, 'lncs')}: {_end(t)} "))
        if c:
            out.append(R(c))
        loc = (f" {v}" if v else "") + (f"({i})" if i else "") + (f", {p}" if p else "") + (f" ({y})" if y else "")
        out.append(R(loc + "."))
        if doi:
            out.append(R(f" {doi}"))
    elif style == "mdpi":
        out.append(R(f"{_authors(ref, 'mdpi')} {_end(t)} "))
        if c:
            out.append(R(c, italic=True))
        if y:
            out.append(R(" "))
            out.append(R(y, bold=True))
        if v:
            out.append(R(", "))
            out.append(R(v, italic=True))
        if p:
            out.append(R(f", {p}"))
        out.append(R("."))
        if doi:
            out.append(R(f" {doi}."))
    elif style == "elsevier-numeric":
        out.append(R(f"{_authors(ref, 'elsevier')}, {t}, "))
        if c:
            out.append(R(c))
        out.append(R((f" {v}" if v else "") + (f" ({y})" if y else "") + (f" {p}" if p else "") + "."))
        if doi:
            out.append(R(f" {doi}."))
    elif style == "mla":
        out.append(R(f"{_end(_authors(ref, 'mla'))} “{_end(t)}” "))
        if c:
            out.append(R(c, italic=True))
        bits = []
        if v: bits.append(f"vol. {v}")
        if i: bits.append(f"no. {i}")
        if y: bits.append(y)
        if p: bits.append(f"pp. {p}")
        out.append(R(", " + ", ".join(bits) + "." if bits else "."))
    else:
        out.append(R(ref.raw))
    out = [r for r in out if r.text]
    out = _tidy(out)
    return _with_label(out, style, n)


def _tidy(runs: List[Run]) -> List[Run]:
    for r in runs:
        r.text = re.sub(r"\s{2,}", " ", r.text)
        r.text = re.sub(r"\.\.(?!\.)", ".", r.text)
        r.text = r.text.replace(" ,", ",").replace(",,", ",").replace(", .", ".")
    return runs


NUMERIC_STYLES = {"ieee", "vancouver", "ama", "lncs", "mdpi", "elsevier-numeric", "springer-vancouver"}


def ref_label(style: str, n: int) -> str:
    if style in ("ieee", "elsevier-numeric"):
        return f"[{n}]"
    if style in ("vancouver", "ama", "lncs", "mdpi", "springer-vancouver"):
        return f"{n}."
    return ""


def _with_label(runs: List[Run], style: str, n: int) -> List[Run]:
    lab = ref_label(style, n)
    if lab:
        return [Run(lab + ("\t" if style in ("ieee", "elsevier-numeric") else " "))] + runs
    return runs


# ------------------------------------------------------------------ in-text citations
BRACKET_RE = re.compile(r"\[(\s*\d{1,3}(?:\s*[-–—,]\s*\d{1,3})*\s*)\]")
PAREN_RE = re.compile(r"\(([^()]*?(?:19|20)\d{2}[a-z]?[^()]*?)\)")
NARR_RE = re.compile(
    r"\b(" + PARTICLE + NAME + r")(?:\s+(?:et\s+al\.?|(?:and|&)\s+(" + PARTICLE + NAME + r")))?\s+\(((?:19|20)\d{2})([a-z])?(?:,\s*(?:p|pp)\.\s*[\d–-]+)?\)"
)


def _expand(spec: str) -> List[int]:
    out = []
    for part in re.split(r"\s*,\s*", spec.strip()):
        m = re.match(r"(\d+)\s*[-–—]\s*(\d+)", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(a, b + 1) if b >= a and b - a < 50 else [a, b])
        elif part.strip().isdigit():
            out.append(int(part))
    return out


def _compress(nums: List[int]) -> str:
    nums = sorted(set(nums))
    parts, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            parts.append(f"{nums[i]}–{nums[j]}")
        else:
            parts.extend(str(x) for x in nums[i:j + 1])
        i = j + 1
    return ", ".join(parts)


def detect_citation_mode(blocks: List[M.Block]) -> str:
    text = " ".join(b.text for b in blocks if b.role in (M.PARA, M.LIST, M.CAPTION))
    nb = len(BRACKET_RE.findall(text))
    na = len(PAREN_RE.findall(text)) + len(NARR_RE.findall(text))
    sup = sum(1 for b in blocks if b.role == M.PARA for r in b.runs
              if r.sup and re.fullmatch(r"\s*\d{1,3}(?:\s*[-–,]\s*\d{1,3})*\s*", r.text or ""))
    if sup > max(nb, na):
        return "superscript"
    return "numeric" if nb >= na and nb > 0 else ("author-year" if na else "none")


def _match_key(ref: Ref):
    return (_fold(ref.first_family), ref.year)


class CitationEngine:
    def __init__(self, refs: List[Ref], target_mode: str, target_style: str):
        self.refs = refs
        self.mode = target_mode
        self.style = target_style
        self.order: List[int] = []       # indices into refs in citation order
        self.unresolved: List[str] = []
        self.by_old = {r.old_num: k for k, r in enumerate(refs) if r.old_num}
        self.by_pos = {k + 1: k for k in range(len(refs))}

    # ---- lookup helpers
    def find_author_year(self, item: str) -> Optional[int]:
        item = re.sub(r"^\s*(?:e\.g\.,?|see|cf\.|i\.e\.,?|also)\s+", "", item.strip(), flags=re.I)
        m = re.search(r"((?:19|20)\d{2})([a-z])?", item)
        if not m:
            return None
        year, suf = m.group(1), m.group(2) or ""
        names = _fold(item[: m.start()])
        names = re.sub(r"et al\.?|,|&|\band\b", " ", names)
        words = [w for w in re.findall(r"[a-z][a-z'’\-]+", names) if len(w) > 1]
        if not words:
            return None
        first = words[0]
        hits = [k for k, r in enumerate(self.refs) if r.year == year and _fold(r.first_family).split()[-1:] and
                (first in _fold(r.first_family).split() or _fold(r.first_family).endswith(first))]
        if len(hits) > 1 and suf:
            hits = [k for k in hits if self.refs[k].suffix == suf] or hits
        if len(hits) > 1 and len(words) > 1:
            hits = [k for k in hits if len(self.refs[k].authors) > 1 and words[1] in _fold(self.refs[k].authors[1][0])] or hits
        return hits[0] if hits else None

    def find_number(self, n: int) -> Optional[int]:
        return self.by_old.get(n, self.by_pos.get(n))

    def _note(self, k: int):
        if k not in self.order:
            self.order.append(k)

    # ---- output builders
    def numeric_text(self, ks: List[int]) -> str:
        for k in ks:
            self._note(k)
        return "PLACEHOLDER"  # resolved later when numbers are final

    def author_year_text(self, ks: List[int], narrative=False) -> str:
        parts = []
        for k in ks:
            r = self.refs[k]
            fam = [a[0] for a in r.authors] or [r.title[:30]]
            amp = " & " if self.mode == "author-year" and self.style == "apa" else " and "
            if len(fam) == 1:
                name = fam[0]
            elif len(fam) == 2:
                name = fam[0] + amp + fam[1]
            else:
                name = fam[0] + " et al."
            if r.etal and len(fam) <= 2:
                name = fam[0] + " et al."
            y = r.year + r.suffix if r.year else "n.d."
            if narrative:
                parts.append(f"{name} ({y})")
            else:
                parts.append(f"{name} {y}" if self.mode == "author-year-nocomma" else f"{name}, {y}")
        if narrative:
            return " and ".join(parts)
        return "(" + "; ".join(parts) + ")"

    # ---- paragraph conversion
    def convert_runs(self, runs: List[Run], source_mode: str) -> List[Run]:
        if source_mode == "none":
            return runs
        text_runs = runs
        segs = _segments(text_runs)
        full = "".join(s[1] for s in segs)
        reps: List[Tuple[int, int, List[Run]]] = []
        numeric_target = self.mode in ("numeric", "superscript")

        if source_mode == "numeric":
            for m in BRACKET_RE.finditer(full):
                ks = [self.find_number(n) for n in _expand(m.group(1))]
                if any(k is None for k in ks):
                    self.unresolved.append(m.group(0))
                    continue
                if numeric_target:
                    for k in ks:
                        self._note(k)
                    reps.append((m.start(), m.end(), [Run("", cite=True, latex=None)]))
                    reps[-1][2][0].text = "\x00" + ",".join(map(str, ks)) + "\x00"
                else:
                    for k in ks:
                        self._note(k)
                    reps.append((m.start(), m.end(), [Run(self.author_year_text(ks))]))
        elif source_mode == "superscript":
            pos = 0
            for kind, t, r in segs:
                if r is not None and r.sup and re.fullmatch(r"\s*\d{1,3}(?:\s*[-–,]\s*\d{1,3})*\s*", t):
                    ks = [self.find_number(n) for n in _expand(t.replace("–", "-"))]
                    if all(k is not None for k in ks):
                        for k in ks:
                            self._note(k)
                        if numeric_target:
                            reps.append((pos, pos + len(t), [Run("\x00" + ",".join(map(str, ks)) + "\x00", cite=True)]))
                        else:
                            reps.append((pos, pos + len(t), [Run(" " + self.author_year_text(ks))]))
                pos += len(t)
        else:  # author-year source
            taken = []
            for m in NARR_RE.finditer(full):
                ks = self.find_author_year(m.group(0).replace("(", " ").replace(")", " "))
                if ks is None:
                    self.unresolved.append(m.group(0))
                    continue
                self._note(ks)
                lead = m.group(0)[: m.group(0).index("(")].rstrip()
                if numeric_target:
                    reps.append((m.start(), m.end(), [Run(lead + " "), Run("\x00" + str(ks) + "\x00", cite=True)]))
                else:
                    reps.append((m.start(), m.end(), [Run(self.author_year_text([ks], narrative=True))]))
                taken.append((m.start(), m.end()))
            for m in PAREN_RE.finditer(full):
                if any(a <= m.start() < b for a, b in taken):
                    continue
                items = [x for x in re.split(r";", m.group(1)) if x.strip()]
                # "Smith, 2019, 2020" -> two items sharing author
                exp = []
                for it in items:
                    ys = re.findall(r"(?:19|20)\d{2}[a-z]?", it)
                    if len(ys) > 1:
                        head = it[: it.find(ys[0])]
                        exp += [head + y for y in ys]
                    else:
                        exp.append(it)
                ks = [self.find_author_year(it) for it in exp]
                if not ks or any(k is None for k in ks):
                    if any(k is not None for k in ks) or re.search(r"[A-Z][a-z]+", m.group(1)):
                        self.unresolved.append(m.group(0))
                    continue
                for k in ks:
                    self._note(k)
                if numeric_target:
                    reps.append((m.start(), m.end(), [Run("\x00" + ",".join(map(str, ks)) + "\x00", cite=True)]))
                else:
                    reps.append((m.start(), m.end(), [Run(self.author_year_text(ks))]))
        if not reps:
            return runs
        return _apply(segs, reps)

    def finalize_numbers(self, blocks: List[M.Block]):
        """Replace placeholders with final numbers (citation order); append uncited refs."""
        for k in range(len(self.refs)):
            if k not in self.order:
                self.order.append(k)
        newnum = {k: i + 1 for i, k in enumerate(self.order)}
        for b in blocks:
            nr = []
            for r in b.runs:
                if r.cite and "\x00" in r.text:
                    ks = [int(x) for x in r.text.strip("\x00").split(",")]
                    nums = [newnum[k] for k in ks]
                    if self.mode == "superscript":
                        nr.append(Run(_compress(nums).replace(" ", ""), sup=True, cite=True))
                    else:
                        nr.append(Run("[" + _compress(nums) + "]", cite=True))
                else:
                    nr.append(r)
            b.runs = nr
            if b.rows:
                pass
        return newnum


def _segments(runs: List[Run]):
    """[(kind, text, run)] where math runs have text '' (non-matchable)."""
    out = []
    for r in runs:
        if r.is_math:
            out.append(("math", "", r))
        else:
            out.append(("text", r.text, r))
    return out


def _apply(segs, reps):
    """Rebuild runs replacing character spans (over concatenated text) with new runs."""
    reps = sorted(reps, key=lambda x: x[0])
    out: List[Run] = []
    pos = 0
    ri = 0
    for kind, t, r in segs:
        if kind == "math":
            out.append(r)
            continue
        start = pos
        end = pos + len(t)
        cur = start
        while ri < len(reps) and reps[ri][0] < end:
            a, b, new = reps[ri]
            if a > cur:
                piece = copy.copy(r); piece.text = t[cur - start: a - start]; out.append(piece)
            if a >= cur:
                for nr in new:
                    nr = copy.copy(nr)
                    nr.bold, nr.italic = (r.bold, r.italic) if not nr.cite else (r.bold, False)
                    out.append(nr)
            if b <= end:
                cur = b
                ri += 1
            else:
                cur = end
                # replacement continues into next run; skip remaining chars there
                reps[ri] = (end, b, [])
                break
        if cur < end:
            piece = copy.copy(r); piece.text = t[cur - start:]; out.append(piece)
        pos = end
    return [x for x in out if x.text or x.is_math]


def sort_author_year(refs: List[Ref]) -> List[int]:
    return sorted(range(len(refs)), key=lambda k: (_fold(refs[k].first_family or refs[k].title),
                                                   _fold(" ".join(a[0] for a in refs[k].authors[1:])),
                                                   refs[k].year, refs[k].suffix))
