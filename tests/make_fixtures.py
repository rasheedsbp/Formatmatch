"""Create realistic test manuscripts (APA-style Word draft) for end-to-end tests."""
import copy
import sys
import io
import os

import docx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))


def fig_png():
    fig, ax = plt.subplots(figsize=(4, 2.5))
    ax.plot([1, 2, 3, 4], [0.71, 0.82, 0.88, 0.93], marker="o")
    ax.set_xlabel("Epoch"); ax.set_ylabel("Accuracy")
    b = io.BytesIO(); fig.savefig(b, format="png", dpi=120); plt.close(fig)
    b.seek(0)
    return b


def make_apa_draft(path):
    d = docx.Document()
    st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
    p = d.add_paragraph(); r = p.add_run("A hybrid deep learning framework for intrusion detection in IoMT networks"); r.bold = True; r.font.size = Pt(16)
    d.add_paragraph("Ravi Kumar¹, Priya Sharma², and John Doe¹")
    d.add_paragraph("¹Department of Computer Science, Osmania University, Hyderabad, India")
    d.add_paragraph("²School of Engineering, University of Leeds, United Kingdom; ravi.k@ou.ac.in")
    d.add_paragraph("Abstract: The Internet of Medical Things (IoMT) is exposed to many attacks. "
                    "This study proposes a hybrid CNN–LSTM model that reaches 98.4% accuracy on CICIoMT2024. "
                    "Results show strong generalisation across attack families.")
    d.add_paragraph("Keywords: IoMT; intrusion detection; deep learning; CNN; LSTM")
    h = d.add_paragraph(); r = h.add_run("1. Introduction"); r.bold = True
    p = d.add_paragraph("Medical devices are increasingly connected (Smith et al., 2020). Prior work has used shallow "
                        "classifiers (Lee & Park, 2019; Brown, 2021), whereas Smith et al. (2020) used CNNs. ")
    r = p.add_run("Escherichia coli"); r.italic = True
    p.add_run(" is not relevant here but tests italics, and H")
    r = p.add_run("2"); r.font.subscript = True
    p.add_run("O tests subscripts.")
    d.add_paragraph("The contributions of this work are:")
    d.add_paragraph("A hybrid CNN–LSTM detector.", style="List Bullet")
    d.add_paragraph("An evaluation on two public datasets.", style="List Bullet")
    h = d.add_paragraph(); r = h.add_run("2. Methodology"); r.bold = True
    h = d.add_paragraph(); r = h.add_run("2.1 Dataset and preprocessing"); r.bold = True
    d.add_paragraph("The dataset statistics are summarised in Table 1 and the loss is defined in Equation (1).")
    # equation copied from pandoc docx
    eqd = docx.Document(os.path.join(HERE, "eq.docx"))
    eq_p = copy.deepcopy(eqd.paragraphs[0]._p)
    d.element.body.insert(len(d.element.body) - 1, eq_p)
    d.add_paragraph("Table 1: Dataset summary")
    t = d.add_table(rows=3, cols=3); t.style = "Table Grid"
    for i, row in enumerate([["Class", "Samples", "Share"], ["Benign", "120,000", "60%"], ["Attack", "80,000", "40%"]]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    h = d.add_paragraph(); r = h.add_run("3. Results and Discussion"); r.bold = True
    d.add_paragraph("Figure 1 shows the training curve; accuracy rises steadily (Kumar, 2022).")
    d.add_picture(fig_png(), width=Inches(4.5))
    d.add_paragraph("Figure 1. Training accuracy over epochs.")
    h = d.add_paragraph(); r = h.add_run("4. Conclusion"); r.bold = True
    d.add_paragraph("The proposed model is accurate and lightweight (Lee & Park, 2019).")
    h = d.add_paragraph(); r = h.add_run("References"); r.bold = True
    for ref in [
        "Brown, T. (2021). Machine learning for network security. IEEE Access, 9, 1123–1135. https://doi.org/10.1109/ACCESS.2021.1234567",
        "Kumar, R. (2022). Lightweight intrusion detection for medical devices. Computers & Security, 115, 102611.",
        "Lee, J., & Park, S. (2019). Shallow classifiers for anomaly detection. Expert Systems with Applications, 120(3), 45–58. https://doi.org/10.1016/j.eswa.2019.01.010",
        "Smith, A. B., Jones, C., & Patel, D. (2020). Deep learning in the Internet of Medical Things. Journal of Network and Computer Applications, 150, 102480. https://doi.org/10.1016/j.jnca.2020.102480",
    ]:
        d.add_paragraph(ref)
    d.save(path)


def make_template(path):
    """A journal-style Word template: 1-col front matter, 2-col body, Arial headings, numbered refs."""
    from docx.enum.section import WD_SECTION
    d = docx.Document()
    st = d.styles["Normal"]; st.font.name = "Cambria"; st.font.size = Pt(9.5)
    sec = d.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.left_margin = sec.right_margin = Inches(0.7)
    def para(text, size=9.5, bold=False, italic=False, align=None, font=None, label=None, label_fmt=None, indent=None):
        p = d.add_paragraph()
        if align: p.alignment = align
        if indent is not None: p.paragraph_format.first_line_indent = Inches(indent)
        if label:
            r = p.add_run(label); r.font.size = Pt(size); r.bold, r.italic = label_fmt
            if font: r.font.name = font
        r = p.add_run(text); r.font.size = Pt(size); r.bold = bold; r.italic = italic
        if font: r.font.name = font
        return p
    para("Template Title Goes Here", 20, True, align=WD_ALIGN_PARAGRAPH.LEFT, font="Arial")
    para("First Author, Second Author", 11, align=WD_ALIGN_PARAGRAPH.LEFT, font="Arial")
    para("Department of X, University of Y, City, Country", 8, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    para("This is the abstract text of the template paper, it should be about two hundred words long and describe the work.", 9,
         label="ABSTRACT: ", label_fmt=(True, False), align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    para("alpha; beta; gamma", 9, label="KEYWORDS: ", label_fmt=(True, False))
    # section break: rest is two columns
    new = d.add_section(WD_SECTION.CONTINUOUS)
    cols = new._sectPr.find(qn("w:cols"))
    if cols is None:
        from docx.oxml import OxmlElement
        cols = OxmlElement("w:cols"); new._sectPr.append(cols)
    cols.set(qn("w:num"), "2"); cols.set(qn("w:space"), "425")
    para("1 INTRODUCTION", 10, True, font="Arial")
    para("Body text of the template is set in Cambria 9.5 pt with a first line indent and justified alignment so that the "
         "renderer can learn it from this exemplar paragraph which is long enough to qualify as body text.", 9.5,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.15)
    para("1.1 Background", 9.5, True, italic=True, font="Arial")
    para("More body text appears here and cites a source [1] and another [2], which shows numbered citations in brackets.",
         9.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.15)
    para("Sample caption text.", 8, label="Figure 1 | ", label_fmt=(True, False), align=WD_ALIGN_PARAGRAPH.LEFT)
    para("Sample table caption.", 8, label="Table 1 | ", label_fmt=(True, False), align=WD_ALIGN_PARAGRAPH.LEFT)
    t = d.add_table(rows=2, cols=2); t.style = "Light Shading"
    para("REFERENCES", 10, True, font="Arial")
    para("[1] A. Author, “Title of paper,” Journal Name, vol. 1, no. 2, pp. 3–4, 2020.", 8)
    para("[2] B. Writer and C. Coder, “Another title,” in Proc. Conf., 2019, pp. 1–9.", 8)
    d.save(path)


if __name__ == "__main__":
    make_apa_draft(os.path.join(HERE, "draft_apa.docx"))
    make_template(os.path.join(HERE, "template_journal.docx"))
    print("ok")


def make_math_algo(path):
    """Manuscript with MathType OLE equations, equation tables, OMML, algorithms (paragraphs + table), complex table."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    from docx.oxml import parse_xml
    from docx.shared import Inches as In
    from PIL import Image, ImageDraw

    d = docx.Document()
    st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(12)
    part = d.part

    def eq_png(text, w=300, h=40):
        im = Image.new("RGB", (w, h), "white"); dr = ImageDraw.Draw(im); dr.text((5, 12), text, fill="black")
        b = io.BytesIO(); im.save(b, "PNG"); return b.getvalue()

    counter = [0]

    def ole_run(text, w_pt=150, h_pt=20):
        counter[0] += 1
        img = Part(PackURI(f"/word/media/mt{counter[0]}.png"), "image/png", eq_png(text), part.package)
        rid_img = part.relate_to(img, RT.IMAGE)
        ole = Part(PackURI(f"/word/embeddings/oleObject{counter[0]}.bin"), "application/vnd.openxmlformats-officedocument.oleObject",
                   b"\xd0\xcf\x11\xe0FAKE-MATHTYPE" + text.encode(), part.package)
        rid_ole = part.relate_to(ole, RT.OLE_OBJECT)
        xml = (f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
               f'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
               f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               f'<w:rPr><w:position w:val="-10"/></w:rPr>'
               f'<w:object w:dxaOrig="{int(w_pt*20)}" w:dyaOrig="{int(h_pt*20)}">'
               f'<v:shape id="_x0000_i10{counter[0]}" type="#_x0000_t75" style="width:{w_pt}pt;height:{h_pt}pt" o:ole="">'
               f'<v:imagedata r:id="{rid_img}" o:title=""/></v:shape>'
               f'<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i10{counter[0]}" DrawAspect="Content" '
               f'ObjectID="_14{counter[0]}" r:id="{rid_ole}"/></w:object></w:r>')
        return parse_xml(xml)

    eqd = docx.Document(os.path.join(HERE, "eq.docx"))
    omml_para = eqd.paragraphs[0]._p
    omml = next(omml_para.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
    inline_omml = next(eqd.paragraphs[1]._p.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))

    p = d.add_paragraph(); r = p.add_run("Energy-Aware Task Scheduling Using Particle Swarm Optimisation"); r.bold = True; r.font.size = Pt(16)
    d.add_paragraph("A. Author, B. Author")
    d.add_paragraph("Department of CSE, Some University, India")
    d.add_paragraph("Abstract: We schedule tasks with PSO and report the energy model used in fog nodes.")
    d.add_paragraph("Keywords: fog computing; PSO; scheduling")
    h = d.add_paragraph(); h.add_run("1. Introduction").bold = True
    p = d.add_paragraph("The energy of node j is ")
    p._p.append(ole_run("E_j = P_j t_j", 60, 14))
    p.add_run(" and the inline OMML term ")
    p._p.append(copy.deepcopy(inline_omml))
    p.add_run(" appears inside the sentence [1].")
    h = d.add_paragraph(); h.add_run("2. System Model").bold = True
    d.add_paragraph("The total energy is given by Eq. (1), using MathType:")
    # MathType display equation with tabs (MTDisplayEquation style)
    p = d.add_paragraph()
    pf = p.paragraph_format
    from docx.enum.text import WD_TAB_ALIGNMENT
    pf.tab_stops.add_tab_stop(In(3.25), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(In(6.5), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run().add_tab()
    p._p.append(ole_run("E = sum_j P_j t_j + E_idle"))
    p.add_run().add_tab()
    p.add_run("(1)")
    d.add_paragraph("The delay constraint is written in an equation table:")
    t = d.add_table(rows=1, cols=3)
    t.cell(0, 0).width = In(0.6); t.cell(0, 1).width = In(5.0); t.cell(0, 2).width = In(0.9)
    t.cell(0, 1).paragraphs[0]._p.append(copy.deepcopy(omml))
    t.cell(0, 2).paragraphs[0].add_run("(2)")
    d.add_paragraph("A plain OMML numbered equation follows.")
    p = d.add_paragraph(); p.add_run().add_tab(); p._p.append(copy.deepcopy(omml)); p.add_run().add_tab(); p.add_run("(3)")
    h = d.add_paragraph(); h.add_run("3. Proposed Method").bold = True
    d.add_paragraph("Algorithm 1 summarises the proposed scheduler, whose update rule uses the velocity term.")
    p = d.add_paragraph(); p.add_run("Algorithm 1: ").bold = True; p.add_run("PSO-based task scheduling")
    lines = [(0, "Input: tasks T, fog nodes F, swarm size N"), (0, "Output: best schedule g"),
             (0, "1: Initialize N particles randomly"), (0, "2: for t = 1 to Tmax do"),
             (1, "3: for each particle i do"), (2, "4: v ← w·v + c1·r1·(p − x)"), (2, "5: x ← x + v"),
             (1, "6: end for"), (0, "7: end for"), (0, "8: return g")]
    for lvl, txt in lines:
        q = d.add_paragraph()
        q.paragraph_format.left_indent = In(0.3 * lvl)
        q.paragraph_format.space_after = Pt(0)
        if lvl == 2 and "v ←" in txt:
            q.add_run("4: v ← ")
            q._p.append(copy.deepcopy(inline_omml))
        else:
            q.add_run(txt)
    d.add_paragraph("After convergence the best particle is mapped to the fog nodes, as discussed in the following "
                    "section where we also compare it with the genetic algorithm baseline in detail for completeness.")
    # Algorithm in a one-cell table with Word auto-numbered lines
    t = d.add_table(rows=1, cols=1); t.style = "Table Grid"
    c = t.cell(0, 0)
    c.paragraphs[0].add_run("Algorithm 2: Genetic algorithm baseline").bold = True
    for txt in ["Generate initial population P", "Evaluate fitness of each chromosome", "Apply crossover and mutation",
                "Return the fittest chromosome"]:
        c.add_paragraph(txt, style="List Number")
    h = d.add_paragraph(); h.add_run("4. Results").bold = True
    d.add_paragraph("Table 1 lists the energy per method with merged headers and a maths cell.")
    d.add_paragraph("Table 1. Energy comparison")
    t = d.add_table(rows=3, cols=3); t.style = "Table Grid"
    a = t.cell(0, 1).merge(t.cell(0, 2)); a.text = "Energy (J)"
    t.cell(0, 0).text = "Method"
    t.cell(1, 0).text = ""; t.cell(1, 1).text = "Mean"; t.cell(1, 2).text = "Std"
    t.cell(2, 0).text = "PSO"; t.cell(2, 1).paragraphs[0]._p.append(copy.deepcopy(inline_omml)); t.cell(2, 2).text = "0.4"
    h = d.add_paragraph(); h.add_run("References").bold = True
    d.add_paragraph("[1] A. Smith, “Fog scheduling,” IEEE Access, vol. 9, pp. 1–10, 2021.")
    d.save(path)


if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "math":
    make_math_algo(os.path.join(HERE, "math_algo.docx"))
    print("math ok")
