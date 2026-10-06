"""Copy verbatim XML (paragraphs, tables, runs) from a source DOCX into the output DOCX.

Everything the XML points to is carried across: images, MathType/OLE objects, hyperlinks, list numbering
definitions and paragraph/character/table styles that the output does not have yet.
"""
from __future__ import annotations

import copy
import posixpath
import re
from typing import Dict, Optional

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from lxml import etree

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TWIPS_PER_CM = 1440 / 2.54


class RawImporter:
    def __init__(self, src_docx, dst_docx):
        self.src = src_docx
        self.dst = dst_docx
        self.rel_map: Dict[str, str] = {}
        self.num_map: Dict[str, str] = {}
        self.copied_styles = set()
        self.counter = 0

    # ---------------------------------------------------------------- relationships
    def _copy_rel(self, rid: str) -> Optional[str]:
        if rid in self.rel_map:
            return self.rel_map[rid]
        src_part = self.src.part
        rel = src_part.rels.get(rid)
        if rel is None:
            return None
        dst_part = self.dst.part
        if rel.is_external:
            new = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            tp = rel.target_part
            self.counter += 1
            base = posixpath.basename(tp.partname)
            stem, ext = posixpath.splitext(base)
            folder = posixpath.dirname(tp.partname)
            name = PackURI(f"{folder}/fm{self.counter}_{stem}{ext}")
            np = Part(name, tp.content_type, tp.blob, dst_part.package)
            new = dst_part.relate_to(np, rel.reltype)
        self.rel_map[rid] = new
        return new

    def _remap_rels(self, el):
        for node in el.iter():
            if not isinstance(node.tag, str):
                continue
            for k, v in list(node.attrib.items()):
                if k.startswith("{" + NS_R + "}") and v:
                    nv = self._copy_rel(v)
                    if nv:
                        node.set(k, nv)
                    else:
                        del node.attrib[k]

    # ---------------------------------------------------------------- numbering
    def _remap_numbering(self, el):
        ids = {n.get(qn("w:val")) for n in el.iter(qn("w:numId")) if n.get(qn("w:val")) not in (None, "0")}
        if not ids:
            return
        try:
            src_num = self.src.part.numbering_part.element
        except Exception:
            return
        dst_num = self.dst.part.numbering_part.element
        for nid in ids:
            if nid in self.num_map:
                continue
            num = src_num.find(f'{qn("w:num")}[@{qn("w:numId")}="{nid}"]')
            if num is None:
                continue
            aid = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            absn = src_num.find(f'{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{aid}"]')
            used_a = [int(a.get(qn("w:abstractNumId"))) for a in dst_num.findall(qn("w:abstractNum"))]
            used_n = [int(n.get(qn("w:numId"))) for n in dst_num.findall(qn("w:num"))]
            new_a = str(max(used_a + [0]) + 1)
            new_n = str(max(used_n + [0]) + 1)
            if absn is not None:
                na = copy.deepcopy(absn)
                na.set(qn("w:abstractNumId"), new_a)
                for x in na.findall(qn("w:nsid")) + na.findall(qn("w:tmpl")):
                    na.remove(x)
                first_num = dst_num.find(qn("w:num"))
                if first_num is not None:
                    first_num.addprevious(na)
                else:
                    tail = dst_num.find(qn("w:numIdMacAtCleanup"))
                    tail.addprevious(na) if tail is not None else dst_num.append(na)
            nn = copy.deepcopy(num)
            nn.set(qn("w:numId"), new_n)
            nn.find(qn("w:abstractNumId")).set(qn("w:val"), new_a if absn is not None else aid)
            tail = dst_num.find(qn("w:numIdMacAtCleanup"))
            tail.addprevious(nn) if tail is not None else dst_num.append(nn)
            self.num_map[nid] = new_n
        for n in el.iter(qn("w:numId")):
            v = n.get(qn("w:val"))
            if v in self.num_map:
                n.set(qn("w:val"), self.num_map[v])

    # ---------------------------------------------------------------- styles
    def _ensure_style(self, sid: str):
        if not sid or sid in self.copied_styles:
            return
        self.copied_styles.add(sid)
        dst_styles = self.dst.styles.element
        if dst_styles.find(f'{qn("w:style")}[@{qn("w:styleId")}="{sid}"]') is not None:
            return
        src = self.src.styles.element.find(f'{qn("w:style")}[@{qn("w:styleId")}="{sid}"]')
        if src is None:
            return
        st = copy.deepcopy(src)
        dst_styles.append(st)
        for tag in ("w:basedOn", "w:link", "w:next"):
            e = st.find(qn(tag))
            if e is not None:
                self._ensure_style(e.get(qn("w:val")))
        self._remap_numbering(st)

    def _styles(self, el):
        for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
            for s in el.iter(qn(tag)):
                self._ensure_style(s.get(qn("w:val")))

    # ---------------------------------------------------------------- public
    def prepare(self, el):
        """Deep-copy `el` and make every reference valid in the destination document."""
        new = copy.deepcopy(el)
        self._remap_rels(new)
        self._remap_numbering(new)
        self._styles(new)
        # drop section breaks and bookmarks/comments that would not resolve
        for sp in list(new.iter(qn("w:sectPr"))):
            sp.getparent().remove(sp)
        for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
            for x in list(new.iter(qn(tag))):
                parent = x.getparent()
                if etree.QName(parent).localname == "r" and len(parent) <= 2:
                    parent.getparent().remove(parent)
                else:
                    parent.remove(x)
        return new


def fit_table(tbl_el, width_cm: float):
    """Scale a fixed-width table so it fits the text column."""
    avail = int(width_cm * TWIPS_PER_CM)
    grid = tbl_el.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol")) if grid is not None else []
    total = sum(int(c.get(qn("w:w"), "0") or 0) for c in cols)
    tblpr = tbl_el.find(qn("w:tblPr"))
    if tblpr is not None:
        ind = tblpr.find(qn("w:tblInd"))
        if ind is not None:
            tblpr.remove(ind)
        tw = tblpr.find(qn("w:tblW"))
        if tw is not None and tw.get(qn("w:type")) == "dxa" and int(tw.get(qn("w:w"), "0") or 0) > avail:
            tw.set(qn("w:w"), str(avail))
    if total <= avail or total == 0:
        return
    f = avail / total
    for c in cols:
        c.set(qn("w:w"), str(int(int(c.get(qn("w:w"), "0") or 0) * f)))
    for tcw in tbl_el.iter(qn("w:tcW")):
        if tcw.get(qn("w:type")) in (None, "dxa"):
            tcw.set(qn("w:w"), str(int(int(tcw.get(qn("w:w"), "0") or 0) * f)))


def fit_equation_tabs(p_el, width_cm: float):
    """Lay out numbered display equations for the new column width: maths centred, number flush right."""
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.SubElement(p_el, qn("w:pPr"))
        p_el.insert(0, ppr)
    ind = ppr.find(qn("w:ind"))
    left = 0
    if ind is not None:
        for k in ("w:left", "w:start"):
            v = ind.get(qn(k))
            if v and v.lstrip("-").isdigit():
                left = int(v)
    w = int(width_cm * TWIPS_PER_CM) - left
    tabs = ppr.find(qn("w:tabs"))
    has_tab_runs = next(p_el.iter(qn("w:tab")), None) is not None and any(
        etree.QName(t.getparent()).localname == "r" for t in p_el.iter(qn("w:tab")))
    text = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()
    numbered = bool(re.fullmatch(r"[\(\[]\s*[\w.\-]+\s*[\)\]]", text))
    if tabs is None or not len(tabs):
        if not (numbered and has_tab_runs):
            return
        # 'TAB maths TAB (n)' without tab stops: add centre + right stops
        tabs = etree.SubElement(ppr, qn("w:tabs"))
        order = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
                 "suppressLineNumbers", "pBdr", "shd", "tabs"]
        ppr.remove(tabs)
        anchor = None
        for ch in ppr:
            if etree.QName(ch).localname not in order:
                anchor = ch
                break
        anchor.addprevious(tabs) if anchor is not None else ppr.append(tabs)
        for val, pos in (("center", w // 2), ("right", w)):
            t = etree.SubElement(tabs, qn("w:tab"))
            t.set(qn("w:val"), val); t.set(qn("w:pos"), str(pos))
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            jc.set(qn("w:val"), "left")
        return
    for t in [t for t in tabs.findall(qn("w:tab")) if t.get(qn("w:val")) != "clear"]:
        pos = int(t.get(qn("w:pos"), "0") or 0)
        if t.get(qn("w:val")) == "center":
            t.set(qn("w:pos"), str(w // 2))
        elif t.get(qn("w:val")) in ("right", "end"):
            t.set(qn("w:pos"), str(w))
        elif pos > w:
            t.set(qn("w:pos"), str(w))
