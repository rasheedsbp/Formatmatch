"""DOCX → Document model (keeps inline formatting, images, tables, OMML equations)."""
from __future__ import annotations

import collections
import re
from typing import List, Optional

import docx
from docx.oxml.ns import qn
from lxml import etree

from . import model as M
from .classify import CAPTION_RE, ParaInfo, RoleClassifier, strip_heading_number
from .runs_util import merge_adjacent, strip_trailing, trim_leading

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_V = "urn:schemas-microsoft-com:vml"
EQ_NUM_RE = re.compile(r"^\s*\(\s*(\d+[a-z]?|[A-Z]?\d+(?:\.\d+)*)\s*\)\s*$")


def _style_chain_attr(style, getter):
    s = style
    while s is not None:
        v = getter(s)
        if v is not None:
            return v
        s = s.base_style
    return None


class DocxReader:
    def __init__(self, path_or_stream):
        self.d = docx.Document(path_or_stream)
        self.part = self.d.part
        self.default_size = self._default_size()
        self.numfmt_cache = {}

    # ---------------- helpers ----------------
    def _default_size(self) -> float:
        try:
            st = self.d.styles["Normal"]
            sz = _style_chain_attr(st, lambda s: s.font.size)
            if sz:
                return sz.pt
        except KeyError:
            pass
        rpr = self.d.styles.element.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr is not None:
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                return int(sz.get(qn("w:val"))) / 2
        return 11.0

    def _para_style(self, p):
        try:
            return p.style
        except Exception:
            return None

    def _style_size(self, style) -> float:
        if style is None:
            return self.default_size
        sz = _style_chain_attr(style, lambda s: s.font.size)
        return sz.pt if sz else self.default_size

    def _style_bold(self, style) -> bool:
        if style is None:
            return False
        return bool(_style_chain_attr(style, lambda s: s.font.bold))

    def _is_ordered(self, numId: str, ilvl: str) -> bool:
        key = (numId, ilvl)
        if key in self.numfmt_cache:
            return self.numfmt_cache[key]
        ordered = False
        try:
            numbering = self.part.numbering_part.element
            num = numbering.find(f'{qn("w:num")}[@{qn("w:numId")}="{numId}"]')
            if num is not None:
                absid = num.find(qn("w:abstractNumId")).get(qn("w:val"))
                absn = numbering.find(f'{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{absid}"]')
                lvl = absn.find(f'{qn("w:lvl")}[@{qn("w:ilvl")}="{ilvl}"]') if absn is not None else None
                if lvl is not None:
                    fmt = lvl.find(qn("w:numFmt"))
                    ordered = fmt is not None and fmt.get(qn("w:val")) not in ("bullet", "none")
        except Exception:
            pass
        self.numfmt_cache[key] = ordered
        return ordered

    # ---------------- run extraction ----------------
    def _runs(self, p_el, para_bold: bool):
        runs: List[M.Run] = []
        images = []

        def rpr_flags(r_el):
            rpr = r_el.find(qn("w:rPr"))
            b = i = sup = sub = False
            b = para_bold
            if rpr is not None:
                be = rpr.find(qn("w:b"))
                if be is not None:
                    b = be.get(qn("w:val")) not in ("0", "false")
                ie = rpr.find(qn("w:i"))
                if ie is not None:
                    i = ie.get(qn("w:val")) not in ("0", "false")
                va = rpr.find(qn("w:vertAlign"))
                if va is not None:
                    sup = va.get(qn("w:val")) == "superscript"
                    sub = va.get(qn("w:val")) == "subscript"
            return b, i, sup, sub

        def walk(el):
            for ch in el:
                tag = etree.QName(ch).localname if isinstance(ch.tag, str) else ""
                ns = etree.QName(ch).namespace if isinstance(ch.tag, str) else ""
                if ns == NS_M and tag in ("oMath", "oMathPara"):
                    runs.append(M.Run(omml=etree.tostring(ch, encoding="unicode")))
                elif tag == "r":
                    b, i, sup, sub = rpr_flags(ch)
                    buf = []
                    for t in ch:
                        tl = etree.QName(t).localname if isinstance(t.tag, str) else ""
                        if tl == "t":
                            buf.append(t.text or "")
                        elif tl == "tab":
                            buf.append(" ")
                        elif tl in ("br", "cr"):
                            buf.append(" ")
                        elif tl == "noBreakHyphen":
                            buf.append("-")
                        elif tl == "sym":
                            buf.append(chr(int(t.get(qn("w:char"), "20"), 16)) if t.get(qn("w:char")) else "")
                        elif tl in ("drawing", "pict", "object"):
                            images.extend(self._images_in(t))
                    if buf:
                        runs.append(M.Run("".join(buf), b, i, sup, sub))
                elif tag in ("hyperlink", "smartTag", "ins", "fldSimple", "customXml", "sdt", "sdtContent", "bdo", "dir"):
                    walk(ch)
                elif tag == "AlternateContent":
                    # prefer the Choice drawing only once
                    for sub in ch:
                        if etree.QName(sub).localname == "Choice":
                            walk(sub)
                            break
        walk(p_el)
        return merge_adjacent(runs), images

    def _images_in(self, el):
        out = []
        for blip in el.iter(f"{{{NS_A}}}blip"):
            rid = blip.get(f"{{{NS_R}}}embed") or blip.get(f"{{{NS_R}}}link")
            if not rid or rid not in self.part.related_parts:
                continue
            ipart = self.part.related_parts[rid]
            ext = ipart.partname.ext if hasattr(ipart.partname, "ext") else "png"
            width = None
            for ext_el in el.iter(f"{{{NS_WP}}}extent"):
                try:
                    width = int(ext_el.get("cx")) / 914400
                except Exception:
                    pass
                break
            out.append((ipart.blob, ext.lower(), width))
        for imd in el.iter(f"{{{NS_V}}}imagedata"):
            rid = imd.get(f"{{{NS_R}}}id")
            if rid and rid in self.part.related_parts:
                ipart = self.part.related_parts[rid]
                out.append((ipart.blob, ipart.partname.ext.lower(), None))
        return out

    # ---------------- main ----------------
    def read(self) -> M.Document:
        doc = M.Document(source_type="docx")
        body = self.d.element.body
        para_objs = {p._p: p for p in self.d.paragraphs}
        items = []
        for el in body.iterchildren():
            tag = etree.QName(el).localname
            if tag == "p":
                items.append(("p", el))
            elif tag == "tbl":
                items.append(("tbl", el))
            elif tag == "sdt":
                content = el.find(qn("w:sdtContent"))
                if content is not None:
                    for sub in content.iterchildren():
                        t2 = etree.QName(sub).localname
                        if t2 in ("p", "tbl"):
                            items.append((t2, sub))

        # body size = dominant run size weighted by text length
        sizes = collections.Counter()
        for p in self.d.paragraphs:
            st = self._para_style(p)
            base = self._style_size(st)
            for r in p.runs:
                sz = r.font.size.pt if r.font.size else base
                sizes[round(sz * 2) / 2] += len(r.text)
        body_size = sizes.most_common(1)[0][0] if sizes else self.default_size
        max_size = max(sizes) if sizes else body_size

        clf = RoleClassifier(body_size=body_size, max_size=max_size)
        from docx.text.paragraph import Paragraph
        from docx.table import Table

        for kind, el in items:
            if kind == "tbl":
                tbl = Table(el, self.d._body)
                rows = []
                for row in tbl.rows:
                    cells = []
                    seen = set()
                    for c in row.cells:
                        if id(c._tc) in seen:
                            cells.append("")
                            continue
                        seen.add(id(c._tc))
                        cells.append(re.sub(r"\s+", " ", c.text).strip())
                    rows.append(cells)
                # 1x1 tables are usually layout boxes (figures / equations)
                if len(rows) == 1 and len(rows[0]) <= 2 and len(tbl._cells) <= 3:
                    for cp in el.iter(qn("w:p")):
                        self._handle_para(Paragraph(cp, self.d._body), clf, doc, body_size)
                    continue
                doc.blocks.append(M.Block(M.TABLE, rows=rows, src=el))
                continue
            p = para_objs.get(el) or Paragraph(el, self.d._body)
            self._handle_para(p, clf, doc, body_size)
        return doc

    def _handle_para(self, p, clf: RoleClassifier, doc: M.Document, body_size: float):
        st = self._para_style(p)
        style_name = st.name if st is not None else ""
        style_bold = self._style_bold(st)
        runs, images = self._runs(p._p, style_bold)

        for blob, ext, width in images:
            doc.blocks.append(M.Block(M.FIGURE, image=blob, image_ext=ext, width_in=width, src=p._p))

        text = "".join(r.text for r in runs if not r.is_math).strip()
        maths = [r for r in runs if r.is_math]

        # display equation: only maths (+ optional "(n)" number)
        if maths and (not text or EQ_NUM_RE.match(text)):
            num = EQ_NUM_RE.match(text).group(1) if text else None
            doc.blocks.append(M.Block(M.EQUATION, runs=maths, eq_num=num, src=p._p))
            return
        if not text and not maths:
            return

        size = self._style_size(st)
        rsz = [r.font.size.pt for r in p.runs if r.font.size and r.text.strip()]
        if rsz:
            size = max(rsz)
        nonempty = [r for r in runs if (r.text.strip() or r.is_math)]
        bold = bool(nonempty) and all(r.bold for r in nonempty if not r.is_math)

        numpr = p._p.find(qn("w:pPr") + "/" + qn("w:numPr")) if p._p.find(qn("w:pPr")) is not None else None
        if numpr is None and st is not None:
            se = st.element.find(qn("w:pPr") + "/" + qn("w:numPr")) if st.element.find(qn("w:pPr")) is not None else None
            numpr = se if (style_name.lower().startswith("list")) else None
        is_list = numpr is not None or style_name.lower().startswith("list")
        ordered, lvl = False, 0
        if numpr is not None:
            nid = numpr.find(qn("w:numId"))
            il = numpr.find(qn("w:ilvl"))
            lvl = int(il.get(qn("w:val"))) if il is not None else 0
            if nid is not None:
                ordered = self._is_ordered(nid.get(qn("w:val")), str(lvl))
        if "number" in style_name.lower():
            ordered = True
        # a numbered heading style uses numPr too -> not a list
        if style_name.lower().startswith("heading"):
            is_list = False

        info = ParaInfo(text=text, style=style_name, size=size, bold=bold, is_list=is_list,
                        list_ordered=ordered, list_level=lvl, is_title_style=style_name.lower() == "title")
        role, level, label, cleaned = clf.feed(info)
        if role is None:
            return
        if role == "abstract_label":
            doc.blocks.append(M.Block(M.ABSTRACT_HEADING, runs=runs, orig=text, src=p._p, src_size=size))
            return
        if cleaned is not None:
            runs = trim_leading(runs, cleaned)
            if not runs and role not in (M.KEYWORDS,):
                return
        runs = strip_trailing(runs)
        blk = M.Block(role, runs=runs, level=level, label=label, style_name=style_name, orig=text,
                      src=p._p, src_size=size, src_bold=bold)
        if role == M.LIST:
            blk.ordered = ordered
            blk.list_level = lvl
            # a list item inside references area etc. handled by classifier
        if role == M.CAPTION:
            m = CAPTION_RE.match(text)
            kind = m.group("kind").lower()
            blk.cap_kind = "table" if kind.startswith("tab") else "figure"
            num = m.group("num")
            blk.cap_num = int(num) if num.isdigit() else None
            blk.runs = trim_leading(runs, m.group("rest")) if m.group("rest").strip() else []
        if role == M.HEADING:
            lvl_num, bare = strip_heading_number(text)
            if lvl_num and not style_name.lower().startswith("heading"):
                blk.level = lvl_num
            blk.runs = trim_leading(runs, bare)
        doc.blocks.append(blk)


def read_docx(path_or_stream) -> M.Document:
    return DocxReader(path_or_stream).read()
