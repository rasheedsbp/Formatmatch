"""PDF → Document model (best-effort: layout-aware text, fonts, images, tables)."""
from __future__ import annotations

import collections
import re
from typing import List

import pymupdf

from . import model as M
from .classify import CAPTION_RE, ParaInfo, RoleClassifier, strip_heading_number
from .runs_util import merge_adjacent, trim_leading

REF_START = re.compile(r"^\s*(\[\d{1,3}\]|\d{1,3}\.\s|\d{1,3}\s(?=[A-Z]))")
BULLET_RE = re.compile(r"^\s*[•▪◦●\uf0b7\uf0a7\u2022\u25aa]\s*")
SPLIT_RE = re.compile(r"^\s*(?:key\s*-?\s*words?|index\s+terms|abstract)\b", re.I)


def _words_table(page, bbox):
    """Rebuild a table from word positions (for rule-only 'booktabs' tables)."""
    words = [w for w in page.get_text("words") if pymupdf.Rect(w[:4]).intersects(bbox)]
    if not words:
        return None
    lines = {}
    for w in sorted(words, key=lambda w: (round(w[3]), w[0])):
        key = next((k for k in lines if abs(k - w[3]) < 3), None)
        lines.setdefault(key if key is not None else w[3], []).append(w)
    rows = []
    for y in sorted(lines):
        ws = sorted(lines[y], key=lambda w: w[0])
        cells, cur = [], [ws[0]]
        for a, b in zip(ws, ws[1:]):
            if b[0] - a[2] > 8:
                cells.append(cur); cur = [b]
            else:
                cur.append(b)
        cells.append(cur)
        rows.append([(c[0][0], c[-1][2], " ".join(x[4] for x in c)) for c in cells])
    ncol = max(len(r) for r in rows)
    if ncol < 2 or len(rows) < 2:
        return None
    anchor = max(rows, key=len)
    centers = [(c[0] + c[1]) / 2 for c in anchor]
    out = []
    for r in rows:
        line = [""] * ncol
        for x0, x1, t in r:
            k = min(range(ncol), key=lambda j: abs(centers[j] - (x0 + x1) / 2))
            line[k] = (line[k] + " " + t).strip()
        out.append(line)
    return out


def _span_flags(span):
    f = span["flags"]
    font = span.get("font", "").lower()
    bold = bool(f & 16) or "bold" in font or "black" in font or "semibold" in font
    italic = bool(f & 2) or "italic" in font or "oblique" in font
    sup = bool(f & 1)
    return bold, italic, sup


def _header_footer_keys(doc):
    """Text lines that repeat at the top/bottom of many pages."""
    cnt = collections.Counter()
    for page in doc:
        h = page.rect.height
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[:5]
            if y1 < h * 0.08 or y0 > h * 0.92:
                key = re.sub(r"\d+", "#", txt.strip())[:80]
                cnt[key] += 1
    n = len(doc)
    return {k for k, v in cnt.items() if n >= 3 and v >= max(2, n * 0.5)}


def _order_blocks(blocks, page_w):
    """Column-aware reading order."""
    mid = page_w / 2
    full, left, right = [], [], []
    for b in blocks:
        x0, x1 = b["bbox"][0], b["bbox"][2]
        if x1 - x0 > page_w * 0.55 or (x0 < mid - 20 and x1 > mid + 20):
            full.append(b)
        elif x1 <= mid + 20:
            left.append(b)
        else:
            right.append(b)
    if len(right) < 2 or len(left) < 2:
        return sorted(blocks, key=lambda b: (round(b["bbox"][1]), b["bbox"][0]))
    # split into horizontal bands by full-width blocks
    full.sort(key=lambda b: b["bbox"][1])
    out = []
    cuts = [b["bbox"][1] for b in full] + [1e9]
    prev = -1e9
    fi = 0
    for cut in cuts:
        band_l = sorted([b for b in left if prev <= b["bbox"][1] < cut], key=lambda b: b["bbox"][1])
        band_r = sorted([b for b in right if prev <= b["bbox"][1] < cut], key=lambda b: b["bbox"][1])
        out += band_l + band_r
        if fi < len(full):
            out.append(full[fi])
            prev = full[fi]["bbox"][3] - 1
            fi += 1
    return out


def read_pdf(path_or_bytes) -> M.Document:
    pdf = pymupdf.open(stream=path_or_bytes, filetype="pdf") if isinstance(path_or_bytes, (bytes, bytearray)) \
        else pymupdf.open(path_or_bytes)
    out = M.Document(source_type="pdf")
    hf = _header_footer_keys(pdf)

    # pass 1: font statistics
    sizes = collections.Counter()
    for page in pdf:
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    sizes[round(s["size"] * 2) / 2] += len(s["text"])
    body_size = sizes.most_common(1)[0][0] if sizes else 10
    max_size = max(sizes) if sizes else body_size
    clf = RoleClassifier(body_size=body_size, max_size=max_size)

    for pno, page in enumerate(pdf):
        pw, ph = page.rect.width, page.rect.height
        # tables
        table_items = []
        try:
            found = page.find_tables().tables
            rebuild = False
            if not found:
                found = page.find_tables(horizontal_strategy="lines", vertical_strategy="text").tables
                rebuild = True
            for t in found:
                rows = _words_table(page, pymupdf.Rect(t.bbox)) if rebuild else \
                    [[(c or "").replace("\n", " ").strip() for c in r] for r in t.extract()]
                if not rows:
                    continue
                if len(rows) >= 2 and max(len(r) for r in rows) >= 2:
                    table_items.append({"bbox": t.bbox, "type": "table", "rows": rows})
        except Exception:
            pass

        def in_table(bb):
            for t in table_items:
                tb = t["bbox"]
                if bb[0] >= tb[0] - 2 and bb[1] >= tb[1] - 2 and bb[2] <= tb[2] + 2 and bb[3] <= tb[3] + 2:
                    return True
            return False

        d = page.get_text("dict")
        blocks = []
        for b in d["blocks"]:
            if b["type"] == 1:
                continue
            if in_table(b["bbox"]):
                continue
            txt = " ".join(s["text"] for l in b.get("lines", []) for s in l["spans"]).strip()
            if not txt:
                continue
            key = re.sub(r"\d+", "#", txt)[:80]
            y0, y1 = b["bbox"][1], b["bbox"][3]
            if (y1 < ph * 0.08 or y0 > ph * 0.92) and (key in hf or re.fullmatch(r"#|page #( of #)?|- # -", key.lower())):
                continue
            blocks.append(b)
        # images (raster)
        for info in page.get_image_info(xrefs=True):
            bb = info["bbox"]
            if (bb[2] - bb[0]) < 40 or (bb[3] - bb[1]) < 40 or not info.get("xref"):
                continue
            blocks.append({"bbox": bb, "type": "image", "xref": info["xref"]})
        blocks += table_items
        for b in _order_blocks(blocks, pw):
            if b.get("type") == "image":
                try:
                    im = pdf.extract_image(b["xref"])
                    out.blocks.append(M.Block(M.FIGURE, image=im["image"], image_ext=im["ext"],
                                              width_in=(b["bbox"][2] - b["bbox"][0]) / 72))
                except Exception:
                    pass
                continue
            if b.get("type") == "table":
                out.blocks.append(M.Block(M.TABLE, rows=b["rows"]))
                continue
            _emit_text_block(b, clf, out, body_size)
    if any(bl.role == M.CAPTION and bl.cap_kind == "figure" for bl in out.blocks) and not out.by_role(M.FIGURE):
        out.warnings.append("PDF figures appear to be vector graphics and could not be extracted; re-insert them in the output.")
    out.warnings.append("PDF input: equations are kept as plain text — check them in the output.")
    return out


def _line_runs(line):
    runs = []
    for s in line["spans"]:
        if not s["text"]:
            continue
        b, i, sup = _span_flags(s)
        runs.append(M.Run(s["text"], bold=b, italic=i, sup=sup and len(s["text"].strip()) <= 4))
    return runs


def _emit_text_block(b, clf: RoleClassifier, out: M.Document, body_size: float):
    lines = b.get("lines", [])
    # Group lines into paragraphs; in references state split per entry.
    groups: List[List[dict]] = [[]]
    min_x = min(l["bbox"][0] for l in lines)
    for l in lines:
        txt = "".join(s["text"] for s in l["spans"]).strip()
        if not txt:
            continue
        if groups[-1] and (BULLET_RE.match(txt) or SPLIT_RE.match(txt)):
            groups.append([])
        elif clf.state == "refs" and groups[-1]:
            prev_txt = "".join(s["text"] for s in groups[-1][-1]["spans"]).strip()
            starts_new = REF_START.match(txt) or (
                abs(l["bbox"][0] - min_x) < 2 and prev_txt.endswith(".") and len(groups[-1]) >= 1
                and any(abs(x["bbox"][0] - min_x) > 6 for x in groups[-1][1:])
            )
            if starts_new:
                groups.append([])
        groups[-1].append(l)
    for g in groups:
        if not g:
            continue
        runs: List[M.Run] = []
        for k, l in enumerate(g):
            lr = _line_runs(l)
            if runs and lr:
                last = runs[-1]
                if last.text.endswith("-") and len(last.text) > 1 and last.text[-2].isalpha() and lr[0].text[:1].islower():
                    last.text = last.text[:-1]
                else:
                    last.text += " "
            runs += lr
        runs = merge_adjacent(runs)
        for r in runs:
            r.text = re.sub(r" {2,}", " ", r.text)
        text = re.sub(r"\s+", " ", "".join(r.text for r in runs)).strip()
        if not text:
            continue
        sizes = collections.Counter()
        for l in g:
            for s in l["spans"]:
                sizes[s["size"]] += len(s["text"])
        size = sizes.most_common(1)[0][0]
        nonempty = [r for r in runs if r.text.strip()]
        bold = all(r.bold for r in nonempty)
        is_list = bool(BULLET_RE.match(text) or re.match(r"^\s*[\-–]\s+", text))
        info = ParaInfo(text=text, size=size, bold=bold, is_list=is_list)
        role, level, label, cleaned = clf.feed(info)
        if role is None:
            continue
        last = out.blocks[-1] if out.blocks else None
        if (last is not None and last.role == M.TITLE and role in (M.AUTHOR, M.OTHER_FRONT, M.AFFIL)
                and abs(size - last.src_size) < 0.6):
            last.runs = merge_adjacent(last.runs + [M.Run(" ")] + runs)
            last.orig += " " + text
            continue
        fonts = collections.Counter()
        for l in g:
            for sp in l["spans"]:
                fonts[sp.get("font", "")] += len(sp["text"])
        font = fonts.most_common(1)[0][0] if fonts else ""
        if role == "abstract_label":
            out.blocks.append(M.Block(M.ABSTRACT_HEADING, runs=runs, orig=text, src_size=size, src_font=font,
                                      src_bold=bold))
            continue
        if cleaned is not None:
            runs = trim_leading(runs, cleaned)
        if is_list:
            runs = trim_leading(runs, BULLET_RE.sub("", re.sub(r"^\s*[\-–]\s+", "", text)))
        blk = M.Block(role, runs=runs, level=level, label=label, orig=text, src_size=size, src_font=font,
                      src_bold=bold)
        if role == M.CAPTION:
            m = CAPTION_RE.match(text)
            blk.cap_kind = "table" if m.group("kind").lower().startswith("tab") else "figure"
            n = m.group("num")
            blk.cap_num = int(n) if n.isdigit() else None
            blk.runs = trim_leading(runs, m.group("rest")) if m.group("rest").strip() else []
        if role == M.HEADING:
            lvl, bare = strip_heading_number(text)
            if lvl:
                blk.level = lvl
            blk.runs = trim_leading(runs, bare)
        out.blocks.append(blk)
