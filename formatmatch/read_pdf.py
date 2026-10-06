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
    bbox = pymupdf.Rect(bbox)
    words = [w for w in page.get_text("words") if bbox.contains(pymupdf.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2))]
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


MATH_FONT_RE = re.compile(r"cmmi|cmsy|cmex|msam|msbm|symbol|math|stix|mt\s?extra|euclid|lmmi|lmsy|lmex|rsfs|esint|"
                          r"txsy|txex|pxsy|pxex|eufm|mtmi|mtsy|newtxmath|ntxmi|ntxsy|cambria", re.I)
EQNUM_SPAN_RE = re.compile(r"^\s*\(\s*\d+[a-z]?(?:\.\d+)?\s*\)\s*$")
OT1 = {"\x0b": "ff", "\x0c": "fi", "\x0d": "fl", "\x0e": "ffi", "\x0f": "ffl", "\x10": "“", "\x11": "”",
       "\x13": "’", "\x14": "‘", "\x15": "–", "\x16": "—", "\x1b": "ff", "\x1c": "fi", "\x1d": "fl"}


def _clean(t: str) -> str:
    if any(ord(c) < 32 for c in t):
        t = "".join(OT1.get(c, "" if ord(c) < 32 and c not in "\t\n" else c) for c in t)
    return t


def _is_math_block(b, body_size) -> bool:
    spans = [s for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
    if not spans:
        return False
    text = " ".join(_clean(s["text"]) for s in spans).strip()
    total = sum(len(s["text"].strip()) for s in spans)
    mathc = sum(len(s["text"].strip()) for s in spans if MATH_FONT_RE.search(s.get("font", "")))
    words = re.findall(r"\b[a-z]{3,}\b", text)
    sentence = len(words) >= 5 or (len(words) >= 3 and text.endswith((".", ":")) and mathc / max(1, total) < 0.3)
    if sentence or len(text) > 220:
        return False
    if EQNUM_SPAN_RE.match(text):
        return True                      # lone "(n)" next to an equation
    if mathc / max(1, total) >= 0.3:
        return True
    if len(text) < 90 and re.search(r"[=∑∫∏≤≥±×∂√∈→←⇒∀∃≈∞]", text) and len(words) <= 2:
        return True
    return False


def _flush_equations(page, pending, out, body_size):
    """Turn consecutive maths-only text blocks into equation images (exact look of the PDF)."""
    if not pending:
        return
    pending = sorted(pending, key=lambda b: (b["bbox"][1], b["bbox"][0]))
    clusters, cur = [], [pending[0]]
    bottom = pending[0]["bbox"][3]
    for b in pending[1:]:
        if b["bbox"][1] - bottom > body_size * 0.9:
            clusters.append(cur); cur = [b]
        else:
            cur.append(b)
        bottom = max(bottom, b["bbox"][3]) if b is not cur[0] else b["bbox"][3]
    clusters.append(cur)
    # an equation number on its own belongs to the cluster at the same height
    merged = []
    for cl in clusters:
        txt = " ".join(_clean(s["text"]) for b in cl for l in b.get("lines", []) for s in l["spans"]).strip()
        if merged and EQNUM_SPAN_RE.match(txt):
            merged[-1] = merged[-1] + cl
        else:
            merged.append(cl)
    clusters = merged
    for cl in clusters:
        rect, num = None, None
        spans = [s for b in cl for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
        xs = [s["bbox"][0] for s in spans]
        mid = (min(xs) + max(s["bbox"][2] for s in spans)) / 2 if spans else 0
        for b in cl:
            if b.get("type") == "mathimg":
                r = pymupdf.Rect(b["bbox"])
                rect = r if rect is None else rect | r
        if spans and rect is not None:
            mid = (min(min(xs), rect.x0) + max(max(s["bbox"][2] for s in spans), rect.x1)) / 2
        for s in spans:
            if EQNUM_SPAN_RE.match(s["text"]) and s["bbox"][0] > mid:
                num = re.sub(r"[()\s]", "", s["text"])
                continue
            r = pymupdf.Rect(s["bbox"])
            rect = r if rect is None else rect | r
        if rect is None:
            continue
        # include drawn parts (fraction bars, radicals) inside the band
        for dr in page.get_drawings():
            dr_r = dr.get("rect")
            if dr_r is not None and dr_r.height < 3 and rect.y0 - 2 <= dr_r.y0 <= rect.y1 + 2 and \
                    dr_r.x0 >= rect.x0 - 20 and dr_r.x1 <= rect.x1 + 20:
                rect |= dr_r
        rect = pymupdf.Rect(rect.x0 - 2, rect.y0 - 2, rect.x1 + 2, rect.y1 + 2)
        pix = page.get_pixmap(clip=rect, dpi=300)
        out.blocks.append(M.Block(M.EQUATION, image=pix.tobytes("png"), image_ext="png", width_in=rect.width / 72,
                                  height_in=rect.height / 72, eq_num=num))


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
    """Column-aware reading order (never drops a block)."""
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
    real_right = [b for b in right if len(" ".join(sp["text"] for l in b.get("lines", []) for sp in l["spans"]).strip()) > 12]
    if len(real_right) < 2 or len(left) < 2:
        return sorted(blocks, key=lambda b: (round(b["bbox"][1]), (b.get("lines") or [b])[0]["bbox"][0]))
    full.sort(key=lambda b: b["bbox"][1])
    out, used = [], set()
    for f in full + [None]:
        cut = f["bbox"][1] if f is not None else 1e9
        band_l = sorted([b for b in left if id(b) not in used and b["bbox"][1] < cut + 1], key=lambda b: b["bbox"][1])
        band_r = sorted([b for b in right if id(b) not in used and b["bbox"][1] < cut + 1], key=lambda b: b["bbox"][1])
        for b in band_l + band_r:
            used.add(id(b))
        out += band_l + band_r
        if f is not None:
            out.append(f)
    return out


def _merge_baselines(blocks, body_size):
    """Join text blocks that continue on the same baseline (text | inline picture | text)."""
    texts = [b for b in blocks if b.get("lines")]
    changed = True
    while changed:
        changed = False
        for a in texts:
            for b in texts:
                if a is b or not a.get("lines") or not b.get("lines"):
                    continue
                fb = b["lines"][0]
                for la in a["lines"]:
                    ya = (la["bbox"][1] + la["bbox"][3]) / 2
                    yb = (fb["bbox"][1] + fb["bbox"][3]) / 2
                    gap = fb["bbox"][0] - la["bbox"][2]
                    if abs(ya - yb) < max(3.0, body_size * 0.45) and -1 <= gap < body_size * 3:
                        la["spans"] = la["spans"] + [dict(sp, text=(" " + sp["text"]) if k == 0 and sp.get("_img") is None else sp["text"])
                                                     for k, sp in enumerate(fb["spans"])]
                        la["bbox"] = (la["bbox"][0], min(la["bbox"][1], fb["bbox"][1]), fb["bbox"][2], max(la["bbox"][3], fb["bbox"][3]))
                        a["lines"] = a["lines"] + b["lines"][1:]
                        a["bbox"] = (min(a["bbox"][0], b["bbox"][0]), min(a["bbox"][1], b["bbox"][1]),
                                     max(a["bbox"][2], b["bbox"][2]), max(a["bbox"][3], b["bbox"][3]))
                        b["lines"] = []
                        changed = True
                        break
                if changed:
                    break
            if changed:
                break
    return [b for b in blocks if b.get("type") in ("image", "table", "mathimg") or b.get("lines")]


def _hits(region, r) -> bool:
    """Rect overlap that also works for zero-height rule lines."""
    R = pymupdf.Rect(region)
    return r.x1 >= R.x0 and r.x0 <= R.x1 and r.y1 >= R.y0 - 1 and r.y0 <= R.y1 + 1


def _ruled_tables(page, region, skip_rect=None):
    """Tables drawn only with horizontal rules (booktabs) inside `region`, excluding algorithm boxes."""
    rules = [d["rect"] for d in page.get_drawings() if d.get("rect") is not None and d["rect"].height < 2.5
             and d["rect"].width > 40 and _hits(region, d["rect"])]
    groups = []
    for r in rules:
        for g in groups:
            if abs(g[0].x0 - r.x0) < 3 and abs(g[0].width - r.width) < 4:
                g.append(r)
                break
        else:
            groups.append([r])
    out = []
    for rs in groups:
        if len(rs) < 2:
            continue
        rs.sort(key=lambda r: r.y0)
        box = pymupdf.Rect(rs[0].x0, rs[0].y0, rs[0].x1, rs[-1].y1 + 1)
        if skip_rect is not None and any(abs(box.x0 - sr.x0) < 3 and abs(box.width - sr.width) < 4 for sr in skip_rect):
            continue
        rows = _words_table(page, box)
        if rows and len(rows) >= 2 and max(len(r) for r in rows) >= 2:
            out.append({"bbox": tuple(box), "type": "table", "rows": rows})
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
        from .algo import is_algo_head
        for strategy in ("default", "text"):
            try:
                found = page.find_tables().tables if strategy == "default" else \
                    page.find_tables(horizontal_strategy="lines", vertical_strategy="text").tables
            except Exception:
                found = []
            for t in found:
                rows = [[(c or "").replace("\n", " ").strip() for c in r] for r in t.extract()] \
                    if strategy == "default" else _words_table(page, pymupdf.Rect(t.bbox))
                if not rows:
                    continue
                if any(is_algo_head(" ".join(c for c in r if c), True) for r in rows[:2]):
                    # ruled algorithm box(es), not a table — but a booktabs table may share the region
                    algo_boxes = []
                    for blk in page.get_text("dict")["blocks"]:
                        for ln in blk.get("lines", []):
                            if is_algo_head(_line_text(ln), True) and pymupdf.Rect(t.bbox).contains(pymupdf.Rect(ln["bbox"])):
                                algo_boxes.append(pymupdf.Rect(ln["bbox"]))
                    groups = []
                    for d in page.get_drawings():
                        r = d.get("rect")
                        if r is not None and r.height < 2.5 and r.width > 40 and _hits(t.bbox, r):
                            for g in groups:
                                if abs(g[0].x0 - r.x0) < 3 and abs(g[0].width - r.width) < 4:
                                    g.append(r)
                                    break
                            else:
                                groups.append([r])
                    skip = []
                    for rs in groups:
                        box = pymupdf.Rect(min(r.x0 for r in rs), min(r.y0 for r in rs), max(r.x1 for r in rs),
                                           max(r.y1 for r in rs))
                        if any(box.x0 - 2 <= a.x0 and box.y0 - 30 <= a.y0 <= box.y1 for a in algo_boxes):
                            skip.append(box)
                    table_items += _ruled_tables(page, t.bbox, skip)
                    continue
                if len(rows) >= 2 and max(len(r) for r in rows) >= 2:
                    table_items.append({"bbox": t.bbox, "type": "table", "rows": rows})
            if table_items:
                break

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
            if y0 > ph * 0.85 and re.fullmatch(r"#|page #( of #)?|- # -|# of #", key.lower()):
                continue
            blocks.append(b)
        # images (raster)
        for info in page.get_image_info(xrefs=True):
            bb = info["bbox"]
            if (bb[2] - bb[0]) >= 15 and 6 <= (bb[3] - bb[1]) < 40 and not in_table(bb):
                blocks.append({"bbox": bb, "type": "mathimg"})      # equation pasted as a picture
                continue
            if (bb[2] - bb[0]) < 40 or (bb[3] - bb[1]) < 40 or not info.get("xref"):
                continue
            blocks.append({"bbox": bb, "type": "image", "xref": info["xref"]})
        # pictures sitting inside a text line are inline maths, not display equations
        for b in [b for b in blocks if b.get("type") == "mathimg"]:
            bb = pymupdf.Rect(b["bbox"])
            neighbours = [t for t in blocks if t.get("lines") and any(
                abs((l["bbox"][1] + l["bbox"][3]) / 2 - (bb.y0 + bb.y1) / 2) < max(4, bb.height / 2)
                and (abs(l["bbox"][2] - bb.x0) < body_size * 3 or abs(bb.x1 - l["bbox"][0]) < body_size * 3)
                for l in t["lines"])]
            if neighbours:
                png = page.get_pixmap(clip=bb, dpi=300).tobytes("png")
                b["type"] = None
                b["lines"] = [{"bbox": tuple(bb), "spans": [{"text": "", "_img": {"blob": png, "ext": "png",
                                                                                  "h": bb.height / 72, "w": bb.width / 72},
                                                            "bbox": tuple(bb), "size": body_size, "flags": 0, "font": ""}]}]
        blocks = _merge_baselines(blocks, body_size)
        blocks += table_items
        pending = []
        for b in _order_blocks(blocks, pw):
            if b.get("type") == "mathimg":
                if clf.state in ("body", "body_pending"):
                    pending.append(b)
                continue
            if b.get("type") not in ("image", "table") and clf.state in ("body", "body_pending"):
                if _is_math_block(b, body_size):
                    pending.append(b)
                    continue
                if pending:
                    t = " ".join(sp["text"] for l in b.get("lines", []) for sp in l["spans"]).strip()
                    y0, y1 = min(x["bbox"][1] for x in pending), max(x["bbox"][3] for x in pending)
                    if len(t) <= 6 and y0 - body_size <= b["bbox"][1] <= y1 + body_size * 0.6:
                        pending.append(b)
                        continue
            if pending:
                _flush_equations(page, pending, out, body_size)
                pending = []
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
            _emit_text_block(b, clf, out, body_size, page)
        _flush_equations(page, pending, out, body_size)
    if any(bl.role == M.CAPTION and bl.cap_kind == "figure" for bl in out.blocks) and not out.by_role(M.FIGURE):
        out.warnings.append("PDF figures appear to be vector graphics and could not be extracted; re-insert them in the output.")
    if out.by_role(M.EQUATION):
        out.warnings.append("PDF input: display equations were captured as high-resolution images from the PDF "
                            "(exact appearance, not editable). Inline maths inside sentences is kept as text.")
    return out


def _line_runs(line):
    runs = []
    for s in line["spans"]:
        if s.get("_img") is not None:
            runs.append(M.Run(img=s["_img"]))
            continue
        if not s["text"]:
            continue
        b, i, sup = _span_flags(s)
        runs.append(M.Run(_clean(s["text"]), bold=b, italic=i, sup=sup and len(s["text"].strip()) <= 4))
    return runs


def _line_text(l):
    return _clean("".join(s["text"] for s in l["spans"])).strip()


def _same_baseline_lines(lines):
    """Join line fragments that sit on the same baseline (e.g. '3:' + 'for each …')."""
    rows = []
    for l in sorted(lines, key=lambda l: (round(l["bbox"][3]), l["bbox"][0])):
        if rows and abs(rows[-1][-1]["bbox"][3] - l["bbox"][3]) < 2.5 and l["bbox"][0] >= rows[-1][-1]["bbox"][0]:
            rows[-1].append(l)
        else:
            rows.append([l])
    return rows


def _algo_take(b, clf, out, body_size, page=None):
    """Algorithm / pseudocode in PDFs: keep every line and its indentation. Returns leftover lines."""
    from .algo import LINENO_RE, is_algo_head, is_algo_line
    lines = [l for l in b.get("lines", []) if _line_text(l)]
    if not lines or clf.state in ("front", "abstract", "refs"):
        return lines
    last = out.blocks[-1] if out.blocks else None
    first = lines[0]
    first_bold = _span_flags(first["spans"][0])[0]
    if is_algo_head(_line_text(first), first_bold):
        title = merge_adjacent(_line_runs(first))
        blk = M.Block(M.ALGORITHM, runs=title, lines=[], orig=_line_text(first), label="open")
        blk._x0 = None
        blk._stop = 1e9
        blk._last_y = first["bbox"][3]
        if page is not None:
            hx0, hx1 = first["bbox"][0] - 30, first["bbox"][0] + (page.rect.width / 2)
            rules = sorted(d["rect"].y0 for d in page.get_drawings()
                           if d.get("rect") is not None and d["rect"].height < 2.5 and d["rect"].width > 60
                           and d["rect"].x0 >= hx0 - 20 and d["rect"].x0 <= first["bbox"][0] + 20
                           and d["rect"].y0 > first["bbox"][1])
            below = [y for y in rules if y > first["bbox"][3] - 1]
            if len(below) >= 2:
                blk._stop = below[1] + 1          # rule under caption, then bottom rule
            elif below and below[0] > first["bbox"][3] + body_size * 2:
                blk._stop = below[0] + 1
        out.blocks.append(blk)
        lines = lines[1:]
        last = blk
    if last is None or last.role != M.ALGORITHM or last.label != "open":
        return lines
    em = max(6.0, body_size)
    rows = _same_baseline_lines(lines)
    for k, row in enumerate(rows):
        y0 = row[0]["bbox"][1]
        txt = " ".join(_line_text(l) for l in row).strip()
        stop = y0 > last._stop or (last._stop == 1e9 and y0 - last._last_y > em * 1.6)
        if stop or CAPTION_RE.match(txt) or (last._stop == 1e9 and not is_algo_line(txt, False, False, False)
                                             and len(txt) > 60):
            last.label = ""
            return [l for r in rows[k:] for l in r]
        # indentation from where the statement text starts (after a line number)
        text_x = row[0]["bbox"][0]
        if len(row) > 1 and LINENO_RE.fullmatch(_line_text(row[0]) + " "):
            text_x = row[1]["bbox"][0]
        elif LINENO_RE.match(_line_text(row[0])):
            text_x = row[0]["bbox"][0]
        if last._x0 is None or text_x < last._x0:
            if last._x0 is not None:
                shift = round((last._x0 - text_x) / (em * 1.2))
                last.lines = [(lv + shift, rs) for lv, rs in last.lines]
            last._x0 = text_x
        lvl = max(0, round((text_x - last._x0) / (em * 1.2)))
        runs = []
        for j, l in enumerate(row):
            if j:
                runs.append(M.Run(" "))
            runs += _line_runs(l)
        last.lines.append((lvl, merge_adjacent(runs)))
        last._last_y = row[-1]["bbox"][3]
    return []


def _emit_text_block(b, clf: RoleClassifier, out: M.Document, body_size: float, page=None):
    left = _algo_take(b, clf, out, body_size, page)
    if not left:
        return
    if len(left) != len([l for l in b.get("lines", []) if _line_text(l)]):
        b = dict(b, lines=left)
    elif out.blocks and out.blocks[-1].role == M.ALGORITHM:
        out.blocks[-1].label = ""
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
