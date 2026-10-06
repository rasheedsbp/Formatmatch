"""Write a prepared Document to DOCX — either by cloning a template's formatting or from the spec."""
from __future__ import annotations

import copy
import io
import re
from typing import List, Optional

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Emu, Inches, Pt, RGBColor
from lxml import etree

from . import model as M
from .model import Run
from .raw_import import RawImporter, fit_equation_tabs, fit_table

ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
PAGE = {"A4": (21.0, 29.7), "Letter": (21.59, 27.94)}
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription"]


PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
             "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
MATH_PR_ORDER = {
    "dPr": ["begChr", "sepChr", "endChr", "grow", "shp", "ctrlPr"],
    "mcPr": ["count", "mcJc"],
    "mPr": ["baseJc", "plcHide", "rSpRule", "cGpRule", "rSp", "cSp", "cGp", "mcs", "ctrlPr"],
    "eqArrPr": ["baseJc", "maxDist", "objDist", "rSpRule", "rSp", "ctrlPr"],
    "naryPr": ["chr", "limLoc", "grow", "subHide", "supHide", "ctrlPr"],
    "fPr": ["type", "ctrlPr"],
}


def _fix_math_order(root):
    """pandoc writes m:dPr children out of schema order; Word is lenient but validators are not."""
    for t in root.iter(f"{{{NS_M}}}t"):
        if t.text and "\u200b" in t.text:
            t.text = t.text.replace("\u200b", "")
    for tag, order in MATH_PR_ORDER.items():
        for pr in root.iter(f"{{{NS_M}}}{tag}"):
            kids = list(pr)
            key = lambda e: order.index(etree.QName(e).localname) if etree.QName(e).localname in order else 99
            if [key(k) for k in kids] != sorted(key(k) for k in kids):
                for k in kids:
                    pr.remove(k)
                for k in sorted(kids, key=key):
                    pr.append(k)


def _put(parent, child, order):
    """Insert/replace `child` in `parent` respecting the schema child order."""
    name = etree.QName(child).localname
    for old in parent.findall(qn("w:" + name)):
        parent.remove(old)
    idx = order.index(name)
    for k, ch in enumerate(parent):
        ln = etree.QName(ch).localname
        if ln in order and order.index(ln) > idx:
            ch.addprevious(child)
            return child
    parent.append(child)
    return child


def _clean_ppr(xml: str):
    el = parse_xml(xml)
    for s in el.findall(qn("w:sectPr")):
        el.remove(s)
    for s in el.findall(qn("w:rPr")):   # paragraph-mark run props: harmless but drop to avoid oddities
        pass
    return el


class DocxRenderer:
    def __init__(self, spec: dict, profile=None):
        self.spec = spec
        self.prof = profile
        self.warnings: List[str] = []
        self.imp: Optional[RawImporter] = None

    # ------------------------------------------------------------ setup
    def _new_document(self):
        if self.prof is not None:
            d = docx.Document(io.BytesIO(self.prof.data))
            body = d.element.body
            for ch in list(body):
                if etree.QName(ch).localname != "sectPr":
                    body.remove(ch)
            return d
        d = docx.Document()
        self._setup_page(d)
        self._setup_styles(d)
        return d

    def _text_width_cm(self) -> float:
        pg = self.spec["page"]
        w = PAGE.get(pg.get("size"), (21.0, 29.7))[0]
        m = pg.get("margins_cm", [2.54] * 4)
        return w - m[2] - m[3]

    def _col_width_cm(self) -> float:
        pg = self.spec["page"]
        n = max(1, int(pg.get("columns", 1)))
        return (self._text_width_cm() - pg.get("col_gap_cm", 0.5) * (n - 1)) / n

    def _setup_page(self, d):
        pg = self.spec["page"]
        s = d.sections[0]
        w, h = PAGE.get(pg.get("size"), (21.0, 29.7))
        s.page_width, s.page_height = Cm(w), Cm(h)
        m = pg.get("margins_cm", [2.54] * 4)
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Cm(m[0]), Cm(m[1]), Cm(m[2]), Cm(m[3])
        sp = s._sectPr
        cols = sp.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            sp.append(cols)
        cols.set(qn("w:num"), str(int(pg.get("columns", 1))))
        cols.set(qn("w:space"), str(int(pg.get("col_gap_cm", 0.5) / 2.54 * 1440)))
        if pg.get("line_numbers"):
            ln = OxmlElement("w:lnNumType")
            ln.set(qn("w:countBy"), "1"); ln.set(qn("w:restart"), "continuous")
            cols.addprevious(ln)
        if pg.get("page_numbers", True):
            fp = s.footer.paragraphs[0]
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._add_field(fp, "PAGE")

    @staticmethod
    def _add_field(p, instr):
        r = p.add_run()
        b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); r._r.append(b)
        r2 = p.add_run()
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r2._r.append(it)
        r3 = p.add_run()
        e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r3._r.append(e)

    def _font_on_style(self, st, family=None, size=None, bold=None, italic=None, smallcaps=None):
        f = st.font
        if family:
            f.name = family
            rpr = st.element.get_or_add_rPr()
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
            for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rf.set(qn(k), family)
            for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rf.get(qn(k)) is not None:
                    del rf.attrib[qn(k)]
        if size:
            f.size = Pt(size)
        if bold is not None:
            f.bold = bold
        if italic is not None:
            f.italic = italic
        if smallcaps is not None:
            f.small_caps = smallcaps
        f.color.rgb = RGBColor(0, 0, 0)

    def _setup_styles(self, d):
        S = self.spec
        fnt = S["font"]
        normal = d.styles["Normal"]
        self._font_on_style(normal, fnt["family"], fnt["size"], False, False)
        pf = normal.paragraph_format
        pf.space_before = Pt(fnt.get("space_before_pt", 0))
        pf.space_after = Pt(fnt.get("space_after_pt", 6))
        pf.line_spacing = fnt.get("line_spacing", 1.0)
        pf.alignment = ALIGN.get(fnt.get("align", "justify"))
        # theme fonts in docDefaults would otherwise override East-Asian etc.
        rpr_def = d.styles.element.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr_def is not None:
            rf = rpr_def.find(qn("w:rFonts"))
            if rf is not None:
                for k in list(rf.attrib):
                    del rf.attrib[k]
                for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                    rf.set(qn(k), fnt["family"])
        levels = S["headings"]["levels"]
        for i in range(4):
            st = d.styles[f"Heading {i + 1}"]
            L = levels[i] if i < len(levels) else levels[-1]
            self._font_on_style(st, fnt["family"], L.get("size", fnt["size"]), L.get("bold", True),
                                L.get("italic", False), L.get("smallcaps", False))
            pf = st.paragraph_format
            pf.alignment = ALIGN.get(L.get("align", "left"))
            pf.space_before = Pt(L.get("space_before_pt", 12))
            pf.space_after = Pt(L.get("space_after_pt", 6))
            pf.keep_with_next = True
            pf.line_spacing = 1.0 if fnt.get("line_spacing", 1) < 1.5 else fnt.get("line_spacing")
            pf.first_line_indent = Cm(0)
        t = S["title"]
        st = d.styles["Title"]
        self._font_on_style(st, fnt["family"], t.get("size", 16), t.get("bold", True), t.get("italic", False))
        ppr = st.element.find(qn("w:pPr"))
        if ppr is not None:
            for bd in ppr.findall(qn("w:pBdr")):
                ppr.remove(bd)
        rpr = st.element.find(qn("w:rPr"))
        if rpr is not None:
            for tag in ("w:spacing", "w:kern"):
                for x in rpr.findall(qn(tag)):
                    rpr.remove(x)
        st.paragraph_format.alignment = ALIGN.get(t.get("align", "center"))
        st.paragraph_format.space_after = Pt(t.get("space_after_pt", 12))
        st.paragraph_format.line_spacing = 1.0
        try:
            cap = d.styles["Caption"]
            c = S["captions"]
            self._font_on_style(cap, fnt["family"], c.get("size", fnt["size"]), False, False)
            cap.paragraph_format.alignment = ALIGN.get(c.get("align", "center"))
            cap.paragraph_format.line_spacing = 1.0
        except KeyError:
            pass

    # ------------------------------------------------------------ primitives
    def _ex(self, *keys) -> Optional[dict]:
        if self.prof is None:
            return None
        for k in keys:
            e = self.prof.exemplars.get(k)
            if e and e.get("pPr") is not None or (e and e.get("rPr")):
                return e
        return None

    def _para(self, ex: Optional[dict] = None, style: Optional[str] = None):
        p = self.d.add_paragraph()
        if ex is not None:
            if ex.get("pPr"):
                old = p._p.find(qn("w:pPr"))
                if old is not None:
                    p._p.remove(old)
                p._p.insert(0, _clean_ppr(ex["pPr"]))
        elif style:
            try:
                p.style = self.d.styles[style]
            except KeyError:
                pass
        return p

    def _fmt(self, p, align=None, size=None, indent_first=None, hanging=None, before=None, after=None,
             line=None, keep_next=None, left=None):
        pf = p.paragraph_format
        if align:
            pf.alignment = ALIGN.get(align)
        if indent_first is not None:
            pf.first_line_indent = Cm(indent_first)
        if hanging:
            pf.left_indent = Cm(hanging)
            pf.first_line_indent = Cm(-hanging)
        if left is not None:
            pf.left_indent = Cm(left)
        if before is not None:
            pf.space_before = Pt(before)
        if after is not None:
            pf.space_after = Pt(after)
        if line is not None:
            pf.line_spacing = line
        if keep_next is not None:
            pf.keep_with_next = keep_next

    def _run(self, p, text, base_rpr: Optional[str] = None, size=None, bold=None, italic=None, sup=False,
             sub=False, smallcaps=None, family=None):
        if text == "":
            return None
        parts = text.split("\n")
        r = None
        for k, part in enumerate(parts):
            if k > 0:
                r.add_break()
            if part == "" and k > 0:
                continue
            tabs = part.split("\t")
            for j, seg in enumerate(tabs):
                if j > 0:
                    r = p.add_run()
                    if base_rpr:
                        r._r.insert(0, parse_xml(base_rpr))
                    r.add_tab()
                if seg == "" and j > 0:
                    continue
                r = p.add_run(seg)
                if base_rpr:
                    old = r._r.find(qn("w:rPr"))
                    if old is not None:
                        r._r.remove(old)
                    r._r.insert(0, parse_xml(base_rpr))
                f = r.font
                if family:
                    f.name = family
                if size:
                    f.size = Pt(size)
                if bold is not None:
                    f.bold = bold
                if italic is not None:
                    f.italic = italic
                if smallcaps is not None:
                    f.small_caps = smallcaps
                if sup:
                    f.superscript = True
                if sub:
                    f.subscript = True
        return r

    def _runs(self, p, runs: List[Run], base_rpr=None, size=None, bold=None, italic=None, smallcaps=None,
              honor_all=False):
        """Write runs; inline bold/italic from the source are kept only where they are partial emphasis."""
        textual = [r for r in runs if not r.is_math and r.text.strip()]
        all_b = bool(textual) and all(r.bold for r in textual)
        all_i = bool(textual) and all(r.italic for r in textual)
        for r in runs:
            if r.is_math:
                self._math(p, r)
                continue
            b = bold
            i = italic
            if honor_all or (r.bold and not all_b):
                b = True if r.bold else b
            if honor_all or (r.italic and not all_i):
                i = True if r.italic else i
            self._run(p, r.text, base_rpr, size=size, bold=b, italic=i, sup=r.sup, sub=r.sub, smallcaps=smallcaps)

    def _insert(self, el):
        body = self.d.element.body
        sp = body.find(qn("w:sectPr"))
        if sp is not None:
            sp.addprevious(el)
        else:
            body.append(el)
        return el

    def _math(self, p, r: Run, inline=True):
        if r.raw is None and r.img is not None and r.omml is None and r.latex is None:
            try:
                p.add_run().add_picture(io.BytesIO(r.img["blob"]), height=Inches(r.img.get("h") or 0.2))
            except Exception:
                self._run(p, "[equation]", italic=True)
            return
        if r.raw is not None:
            if self.imp is not None:
                p._p.append(self.imp.prepare(r.raw))
            elif r.img and r.img.get("blob"):
                try:
                    p.add_run().add_picture(io.BytesIO(r.img["blob"]), height=Inches(r.img.get("h") or 0.2))
                except Exception:
                    self._run(p, "[equation]", italic=True)
            return
        if r.omml:
            el = parse_xml(r.omml)
            if etree.QName(el).localname == "oMathPara" and inline:
                for om in el.iter(f"{{{NS_M}}}oMath"):
                    p._p.append(copy.deepcopy(om))
            else:
                p._p.append(el)
        elif r.latex:
            self._run(p, r.latex, italic=True)

    # ------------------------------------------------------------ block writers
    def _front_para(self, b: M.Block, key: str, spec_key: str, style=None):
        S = self.spec
        ex = self._ex(key, "author" if key == "affiliation" else key)
        p = self._para(ex, style)
        if ex:
            self._runs(p, b.runs, ex.get("rPr"))
        else:
            c = S[spec_key]
            self._fmt(p, align=c.get("align", "center"), indent_first=0, after=c.get("space_after_pt", 4),
                      line=1.0)
            self._runs(p, b.runs, size=c.get("size"), bold=c.get("bold"), italic=c.get("italic"))
        return p

    def _abstract(self, blocks: List[M.Block]):
        S = self.spec["abstract"]
        fsz = S.get("size")
        ex = self._ex("abstract")
        layout = S.get("layout", "block")
        if layout == "block":
            hx = self._ex("abstract_heading")
            hp = self._para(hx)
            if hx:
                self._run(hp, S.get("heading", "Abstract"), hx.get("rPr"))
            else:
                self._fmt(hp, align="center" if self.spec["title"].get("align") == "center" else "left",
                          indent_first=0, before=6, after=3, keep_next=True)
                self._run(hp, S.get("heading", "Abstract"), size=fsz, bold=S.get("label_bold", True),
                          italic=S.get("label_italic", False))
        for k, b in enumerate(blocks):
            p = self._para(ex)
            if not ex:
                self._fmt(p, align=S.get("align", "justify"), indent_first=0, after=3, line=None,
                          left=S.get("indent_cm") or None)
            lab_rpr = ex.get("label_rPr") if ex else None
            body_rpr = ex.get("rPr") if ex else None
            if layout == "inline" and k == 0:
                lab = S.get("heading", "Abstract") + S.get("suffix", "—")
                self._run(p, lab, lab_rpr, size=None if lab_rpr else fsz,
                          bold=None if lab_rpr else S.get("label_bold", True),
                          italic=None if lab_rpr else S.get("label_italic", False))
            if b.label:
                self._run(p, b.label + " – ", lab_rpr or body_rpr, size=None if (lab_rpr or body_rpr) else fsz, bold=True)
            self._runs(p, b.runs, body_rpr, size=None if body_rpr else fsz,
                       bold=None if body_rpr else (S.get("body_bold") or None),
                       italic=None if body_rpr else (S.get("body_italic") or None))

    def _keywords(self, b: M.Block):
        K = self.spec["keywords"]
        ex = self._ex("keywords")
        p = self._para(ex)
        if not ex:
            self._fmt(p, align=self.spec["abstract"].get("align", "justify"), indent_first=0, before=3, after=6)
        lab_rpr = ex.get("label_rPr") if ex else None
        body_rpr = ex.get("rPr") if ex else None
        self._run(p, K.get("label", "Keywords") + K.get("suffix", ": "), lab_rpr,
                  size=None if lab_rpr else K.get("size"), bold=None if lab_rpr else K.get("label_bold", True),
                  italic=None if lab_rpr else K.get("label_italic", False))
        self._runs(p, b.runs, body_rpr, size=None if body_rpr else K.get("size"),
                   italic=None if body_rpr else (K.get("body_italic") or None),
                   bold=None if body_rpr else (K.get("body_bold") or None))

    def _heading(self, b: M.Block):
        lvl = max(1, min(b.level, 4))
        ex = self._ex(*[f"h{x}" for x in range(lvl, 0, -1)]) if self.prof else None
        if ex and self.prof.exemplars.get(f"h{lvl}") is None:
            ex = None    # don't reuse a higher-level look for a lower level; fall back to spec
        if ex:
            p = self._para(ex)
            self._run(p, b.prefix, ex.get("rPr"))
            self._runs(p, b.runs, ex.get("rPr"))
            return
        p = self.d.add_paragraph(style=f"Heading {lvl}")
        L = self.spec["headings"]["levels"][lvl - 1]
        self._run(p, b.prefix)
        self._runs(p, b.runs)

    def _body_para(self, b: M.Block, after_heading: bool):
        F = self.spec["font"]
        ex = self._ex("para_first", "para", "para_short") if after_heading else self._ex("para", "para_short")
        p = self._para(ex)
        if not ex:
            ind = F.get("first_line_indent_cm", 0)
            # APA/IEEE: no indent right after a heading is common; keep spec indent otherwise
            self._fmt(p, indent_first=0 if (after_heading and ind and self.spec["headings"]["scheme"] != "none") else ind)
        self._runs(p, b.runs, ex.get("rPr") if ex else None)

    def _list(self, b: M.Block):
        key = "list_number" if b.ordered else "list_bullet"
        ex = self._ex(key) if self.prof else None
        if ex:
            p = self._para(ex)
            if not ex.get("auto_num"):
                self._run(p, ("• " if not b.ordered else "– "), ex.get("rPr"))
            self._runs(p, b.runs, ex.get("rPr"))
            return
        style = "List Number" if b.ordered else "List Bullet"
        if self.prof is not None:
            try:
                self.d.styles[style]
            except KeyError:
                style = None
        p = self._para(None, style)
        if style is None:
            self._fmt(p, left=0.63, indent_first=-0.4)
            self._run(p, "• ")
        self._runs(p, b.runs)

    def _caption(self, b: M.Block):
        C = self.spec["captions"]
        key = f"caption_{b.cap_kind}"
        ex = self._ex(key, "caption_figure", "caption_table")
        p = self._para(ex, "Caption")
        lab_rpr = ex.get("label_rPr") if ex else None
        body_rpr = ex.get("rPr") if ex else None
        auto = bool(ex and ex.get("auto_num") and self.prof.exemplars.get(key))
        if not ex:
            self._fmt(p, align=C.get("align", "center"), indent_first=0, before=4, after=6, line=1.0,
                      keep_next=(b.cap_kind == "table" and C.get("table_position", "above") == "above"))
        if not auto:
            pre = b.prefix
            label_part, sep = (pre.split("\n")[0], "\n") if "\n" in pre else (pre.rstrip(), pre[len(pre.rstrip()):])
            self._run(p, label_part, lab_rpr, size=None if lab_rpr else C.get("size"),
                      bold=None if lab_rpr else C.get("label_bold", False),
                      italic=None if lab_rpr else C.get("label_italic", False),
                      smallcaps=True if (b.cap_kind == "table" and C.get("table_label_newline") and not lab_rpr) else None)
            if sep == "\n":
                last = p.runs[-1] if p.runs else self._run(p, "")
                if last is not None:
                    last.add_break()
            else:
                self._run(p, sep, body_rpr or lab_rpr, size=None if (body_rpr or lab_rpr) else C.get("size"))
        sc = True if (b.cap_kind == "table" and C.get("table_caption_case") == "upper" and not body_rpr) else None
        self._runs(p, b.runs, body_rpr, size=None if body_rpr else C.get("size"),
                   italic=None if body_rpr else (C.get("body_italic") or None), smallcaps=sc)

    def _figure(self, b: M.Block):
        ex = self._ex("figure")
        p = self._para(ex)
        if not ex:
            self._fmt(p, align="center", indent_first=0, before=6, after=3, keep_next=True, line=1.0)
        maxw = self._col_width_cm() / 2.54
        w = min(b.width_in or maxw, maxw)
        data = b.image
        try:
            p.add_run().add_picture(io.BytesIO(data), width=Inches(w))
        except Exception:
            try:
                from PIL import Image
                im = Image.open(io.BytesIO(data))
                buf = io.BytesIO(); im.convert("RGB").save(buf, format="PNG"); buf.seek(0)
                p.add_run().add_picture(buf, width=Inches(w))
            except Exception:
                self._run(p, "[Figure could not be converted — insert the original image here]", italic=True)
                self.warnings.append("A figure in an unsupported image format (e.g. EMF) must be re-inserted manually.")

    def _table(self, b: M.Block):
        T = self.spec["tables"]
        if b.raw and self.imp is not None:
            # complex table (maths, merged cells, images): keep it exactly, only fit it to the column
            for el in b.raw:
                new = self.imp.prepare(el)
                fit_table(new, self._col_width_cm())
                self._insert(new)
            sp = self.d.add_paragraph()
            sp.paragraph_format.space_after = Pt(2)
            sp.paragraph_format.line_spacing = 0.6
            return
        rows = b.rows or [[""]]
        ncol = max(len(r) for r in rows)
        t = self.d.add_table(rows=len(rows), cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        if self.prof is not None and self.prof.table_tblPr:
            old = t._tbl.find(qn("w:tblPr"))
            new = parse_xml(self.prof.table_tblPr)
            t._tbl.replace(old, new)
        else:
            self._table_borders(t, T.get("borders", "booktabs"))
        tblPr = t._tbl.tblPr
        tw = OxmlElement("w:tblW")
        tw.set(qn("w:type"), "pct"); tw.set(qn("w:w"), "5000")
        _put(tblPr, tw, TBLPR_ORDER)
        for i, row in enumerate(rows):
            for j in range(ncol):
                cell = t.cell(i, j)
                txt = row[j] if j < len(row) else ""
                cp = cell.paragraphs[0]
                cp.paragraph_format.first_line_indent = Cm(0)
                cp.paragraph_format.space_before = Pt(1)
                cp.paragraph_format.space_after = Pt(1)
                cp.paragraph_format.line_spacing = 1.0
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER if (i == 0 or re.fullmatch(r"[\d.,%±\-–+()\s]+", txt or "x")) else WD_ALIGN_PARAGRAPH.LEFT
                cell_runs = b.rows_runs[i][j] if (b.rows_runs and i < len(b.rows_runs) and j < len(b.rows_runs[i])) else None
                if cell_runs and any(r.is_math for r in cell_runs):
                    self._runs(cp, cell_runs, size=T.get("size"), bold=(i == 0 and T.get("header_bold", True)) or None,
                               honor_all=True)
                else:
                    self._run(cp, txt, size=T.get("size"), bold=(i == 0 and T.get("header_bold", True)) or None)
        if T.get("borders", "booktabs") == "booktabs" and not (self.prof and self.prof.table_tblPr):
            for c in t.rows[0].cells:
                tcPr = c._tc.get_or_add_tcPr()
                bd = OxmlElement("w:tcBorders")
                bt = OxmlElement("w:bottom")
                bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "6"); bt.set(qn("w:color"), "000000")
                bd.append(bt); tcPr.append(bd)
        # small spacer after table
        sp = self.d.add_paragraph()
        sp.paragraph_format.space_after = Pt(2)
        sp.paragraph_format.line_spacing = 0.6

    @staticmethod
    def _table_borders(t, kind):
        tblPr = t._tbl.tblPr
        bd = OxmlElement("w:tblBorders")
        spec = {"top": 12, "bottom": 12} if kind == "booktabs" else \
            {k: 4 for k in ("top", "left", "bottom", "right", "insideH", "insideV")}
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{edge}")
            if edge in spec:
                e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(spec[edge])); e.set(qn("w:color"), "000000")
            else:
                e.set(qn("w:val"), "nil")
            bd.append(e)
        _put(tblPr, bd, TBLPR_ORDER)

    def _equation(self, b: M.Block):
        if b.raw and self.imp is not None:
            for el in b.raw:
                new = self.imp.prepare(el)
                if etree.QName(new).localname == "tbl":
                    fit_table(new, self._col_width_cm())
                    for tc in new.iter(qn("w:tc")):           # number aligned with the equation
                        tcpr = tc.find(qn("w:tcPr"))
                        if tcpr is None:
                            tcpr = OxmlElement("w:tcPr"); tc.insert(0, tcpr)
                        for old in tcpr.findall(qn("w:vAlign")):
                            tcpr.remove(old)
                        va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "center")
                        tcpr.append(va)
                else:
                    fit_equation_tabs(new, self._col_width_cm())
                self._insert(new)
            return
        ex = self._ex("equation")
        p = self._para(ex)
        if not ex:
            self._fmt(p, align="left" if b.eq_num else "center", indent_first=0, before=4, after=4, line=1.0)
        if b.image and not b.runs:
            # equation captured from a PDF as a picture: keep its exact size (shrink only if wider than the column)
            col_in = self._col_width_cm() / 2.54 * (0.82 if b.eq_num else 1.0)
            w = min(b.width_in or col_in, col_in)
            pic_run = Run(img={"blob": b.image, "ext": b.image_ext, "h": (b.height_in or 0.3) * w / (b.width_in or w)})
            b = M.Block(M.EQUATION, runs=[pic_run], eq_num=b.eq_num)
        if b.eq_num:
            w = self._col_width_cm()
            ind = p.paragraph_format.left_indent.cm if p.paragraph_format.left_indent else 0
            ts = p.paragraph_format.tab_stops
            ts.add_tab_stop(Cm(w / 2 - ind), WD_TAB_ALIGNMENT.CENTER)
            ts.add_tab_stop(Cm(w - ind), WD_TAB_ALIGNMENT.RIGHT)
            p.paragraph_format.first_line_indent = Cm(0)
            p.add_run().add_tab()
            for r in b.runs:
                self._math(p, r, inline=True)
            p.add_run().add_tab()
            nr = self._run(p, f"({b.eq_num})")
            pics = [r for r in b.runs if r.img is not None and r.raw is None and r.omml is None]
            if pics and nr is not None:
                h_pt = max((r.img.get("h") or 0.2) for r in pics) * 72
                rpr = nr._r.get_or_add_rPr()
                pos = OxmlElement("w:position")
                pos.set(qn("w:val"), str(int(max(0, h_pt / 2 - 3) * 2)))
                rpr.append(pos)
        else:
            for r in b.runs:
                self._math(p, r, inline=False)

    def _algorithm(self, b: M.Block):
        if b.raw and self.imp is not None:
            news = [self.imp.prepare(el) for el in b.raw]
            paras = [n for n in news if etree.QName(n).localname == "p"]
            for n in news:
                if etree.QName(n).localname == "tbl":
                    fit_table(n, self._col_width_cm())
                self._insert(n)
            if etree.QName(news[-1]).localname == "tbl":
                sp = self.d.add_paragraph()
                sp.paragraph_format.space_after = Pt(4)
                sp.paragraph_format.line_spacing = 0.6
            if 1 < len(paras) <= 45:                      # keep the algorithm on one page/column
                for pe in paras[:-1]:
                    ppr = pe.find(qn("w:pPr"))
                    if ppr is None:
                        ppr = OxmlElement("w:pPr"); pe.insert(0, ppr)
                    if ppr.find(qn("w:keepNext")) is None:
                        kn = OxmlElement("w:keepNext")
                        anchor = ppr.find(qn("w:pStyle"))
                        if anchor is not None:
                            anchor.addnext(kn)
                        else:
                            ppr.insert(0, kn)
            return
        # rebuilt (PDF / LaTeX sources): ruled algorithm box
        F = self.spec["font"]
        size = max(7.0, F["size"] - 1)

        def rule(p, edge):
            ppr = p._p.get_or_add_pPr()
            bdr = ppr.find(qn("w:pBdr"))
            if bdr is None:
                bdr = OxmlElement("w:pBdr")
                _put(ppr, bdr, PPR_ORDER)
            e = OxmlElement(f"w:{edge}")
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "8"); e.set(qn("w:space"), "1"); e.set(qn("w:color"), "000000")
            bdr.append(e)

        head = self.d.add_paragraph()
        self._fmt(head, align="left", indent_first=0, before=6, after=2, line=1.0, keep_next=True)
        if b.runs:
            self._runs(head, b.runs, size=size, honor_all=True)
        else:
            self._run(head, b.orig or "Algorithm", size=size, bold=True)
        rule(head, "top"); rule(head, "bottom")
        n = len(b.lines or [])
        for k, (lvl, runs) in enumerate(b.lines or []):
            p = self.d.add_paragraph()
            self._fmt(p, align="left", indent_first=0, before=0, after=0, line=1.0, keep_next=k < n - 1 and n <= 45,
                      left=0.2 + 0.5 * float(lvl))
            self._runs(p, runs, size=size, honor_all=True)
            if k == n - 1:
                rule(p, "bottom")
        sp = self.d.add_paragraph()
        sp.paragraph_format.space_after = Pt(4)
        sp.paragraph_format.line_spacing = 0.6

    def _reference(self, b: M.Block):
        R = self.spec["references"]
        ex = self._ex("reference")
        p = self._para(ex)
        runs = b.runs
        if ex and ex.get("auto_num") and runs and re.match(r"^\s*(\[\d+\]|\d+\.)\s*", runs[0].text):
            runs = [Run(re.sub(r"^\s*(\[\d+\]|\d+\.)\s*", "", runs[0].text))] + runs[1:]
        if ex and runs and "\t" in runs[0].text and not (ex.get("pPr") and "hanging" in ex["pPr"]):
            runs = [Run(runs[0].text.replace("\t", " "))] + runs[1:]
        if not ex:
            h = R.get("hanging_indent_cm", 0.63)
            self._fmt(p, align="justify" if R.get("style") == "ieee" else "left", hanging=h, after=2,
                      line=1.0 if self.spec["font"].get("line_spacing", 1) < 2 else None)
            if runs and "\t" in runs[0].text:
                p.paragraph_format.tab_stops.add_tab_stop(Cm(h))
        self._runs(p, runs, ex.get("rPr") if ex else None, size=None if ex else R.get("size"), honor_all=True)

    # ------------------------------------------------------------ main
    def render(self, doc: M.Document) -> bytes:
        self.d = self._new_document()
        if doc.src_docx is not None:
            self.imp = RawImporter(doc.src_docx, self.d)
        S = self.spec
        blocks = doc.blocks
        front_roles = {M.TITLE, M.AUTHOR, M.AFFIL, M.OTHER_FRONT, M.ABSTRACT, M.KEYWORDS}
        # index where the body starts (first block not in front matter)
        body_start = next((k for k, b in enumerate(blocks) if b.role not in front_roles), len(blocks))
        i = 0
        prev_role = None
        while i < len(blocks):
            b = blocks[i]
            if i == body_start:
                self._section_break_after_front()
            r = b.role
            if r == M.TITLE:
                self._front_para(b, "title", "title", style="Title" if self.prof is None else None)
            elif r == M.AUTHOR:
                self._front_para(b, "author", "authors")
            elif r == M.AFFIL:
                self._front_para(b, "affiliation", "affiliation")
            elif r == M.OTHER_FRONT:
                self._front_para(b, "front", "affiliation")
            elif r == M.ABSTRACT:
                group = []
                while i < len(blocks) and blocks[i].role == M.ABSTRACT:
                    group.append(blocks[i]); i += 1
                self._abstract(group)
                prev_role = M.ABSTRACT
                continue
            elif r == M.KEYWORDS:
                self._keywords(b)
            elif r == M.HEADING:
                self._heading(b)
            elif r == M.PARA:
                self._body_para(b, after_heading=prev_role == M.HEADING)
            elif r == M.LIST:
                self._list(b)
            elif r == M.CAPTION:
                self._caption(b)
            elif r == M.FIGURE:
                self._figure(b)
            elif r == M.TABLE:
                self._table(b)
            elif r == M.EQUATION:
                self._equation(b)
            elif r == M.ALGORITHM:
                self._algorithm(b)
            elif r == M.REFERENCE:
                self._reference(b)
            prev_role = r
            i += 1
        if body_start >= len(blocks):
            pass
        _fix_math_order(self.d.element.body)
        zoom = self.d.settings.element.find(qn("w:zoom"))
        if zoom is not None and zoom.get(qn("w:percent")) is None:
            zoom.set(qn("w:percent"), "100")
        buf = io.BytesIO()
        self.d.save(buf)
        return buf.getvalue()

    def _section_break_after_front(self):
        pg = self.spec["page"]
        final = self.d.element.body.find(qn("w:sectPr"))
        if self.prof is not None:
            if self.prof.front_section:
                p = self.d.add_paragraph()
                sp = parse_xml(self.prof.front_section)
                p._p.get_or_add_pPr().append(sp)
                p.paragraph_format.space_after = Pt(0)
            return
        if int(pg.get("columns", 1)) > 1 and pg.get("front_matter_single_column", True):
            p = self.d.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            sp = copy.deepcopy(final)
            for hf in sp.findall(qn("w:headerReference")) + sp.findall(qn("w:footerReference")):
                sp.remove(hf)
            cols = sp.find(qn("w:cols"))
            cols.set(qn("w:num"), "1")
            p._p.get_or_add_pPr().append(sp)
            # final section continues on the same page
            ty = final.find(qn("w:type"))
            if ty is None:
                ty = OxmlElement("w:type")
                pgsz = final.find(qn("w:pgSz"))
                pgsz.addprevious(ty)
            ty.set(qn("w:val"), "continuous")


def render_docx(doc: M.Document, spec: dict, profile=None):
    r = DocxRenderer(spec, profile)
    data = r.render(doc)
    return data, r.warnings
