# FormatMatch — Manuscript Formatter

Arrange any manuscript into a target format: upload the paper, give the format as a **Word template / sample
paper**, **author guidelines**, or a **journal name** (any combination), and download the paper rebuilt in that
format as **Word (.docx)** and a **LaTeX project (.zip)**, plus a compliance report.

By CS Software Solutions — cssoftwaresolutions.in · info@cssoftwaresolutions.tech

## What it does

| Area | Behaviour |
|---|---|
| Inputs | .docx, .doc/.odt/.rtf (via LibreOffice), .pdf (layout-aware, 1- or 2-column), .tex or .zip LaTeX project with .bib (via pandoc), Markdown |
| Format sources | Word template/.dotx (styles cloned exactly), published sample PDF (layout measured), guidelines text/file/URL (rule-based or optional Groq AI), journal name (Crossref publisher lookup → house style) |
| House styles | IEEE journal/conference, Elsevier (numbered/Harvard), Springer Nature (author–year/numbered), Springer LNCS, MDPI, Emerald, Taylor & Francis, Wiley, APA 7, PLOS, Frontiers, generic |
| Equations & algorithms | Never reflowed: Word equations (OMML), MathType / Equation Editor objects, equation tables, algorithm/pseudocode blocks (plain lines, tables, Word-numbered) and complex tables (merged cells, maths) are copied verbatim with their images, OLE objects, numbering and styles; only fitted to the column width. LaTeX input keeps the original `algorithm`/`equation`/`align` source. PDF input: display equations captured as 300-dpi images, algorithm lines and indentation kept |
| Structure | Title, authors, affiliations, abstract (inline/block/structured), keywords, headings (1. / 1.1, I. / A. / 1), unnumbered), lists, tables, figures, OMML equations, captions |
| Citations | Converts author–year ⇄ numbered ⇄ superscript, renumbers by first citation, compresses ranges [2–5], flags unmatched citations and uncited references |
| References | Parses APA, Harvard, IEEE, Vancouver, Chicago, LNCS, MDPI, Springer, Emerald entries; re-writes in 13 styles; optional Crossref verification (fills DOI/volume/pages) |
| Captions | Renumbers figures/tables in order, applies labels (Fig./Figure, TABLE I), positions (above/below), updates in-text cross-references |
| Report | Abstract/keyword/title/word limits, required sections, captions present and cited, every change made, items needing attention |

## Run locally

```bash
pip install -r requirements.txt
# optional but recommended: pandoc (LaTeX input + equation export) and LibreOffice (.doc input, previews)
streamlit run app.py
```

## Deploy on Streamlit Cloud

Push this folder to GitHub and create an app pointing at `app.py`. `packages.txt` installs pandoc and LibreOffice
Writer. For AI guideline reading add `GROQ_API_KEY` under *Secrets* (optional — rule-based reading works without it).

## Use from Python

```python
from formatmatch.engine import FormatSource, build_spec, format_manuscript, learn_template

ov, profile, notes = learn_template("journal_template.docx", open("journal_template.docx", "rb").read())
src = FormatSource(preset="ieee_journal", template_overrides=ov, profile=profile)
res = format_manuscript("paper.docx", open("paper.docx", "rb").read(), build_spec(src), profile)
open("paper_formatted.docx", "wb").write(res.docx)
open("paper_latex.zip", "wb").write(res.latex_zip)
print(res.report_md)
```

## Layout

```
app.py                     Streamlit UI
formatmatch/
  read_docx.py read_pdf.py read_any.py   manuscript → document model
  classify.py                            role detection (title, abstract, headings, captions, refs…)
  spec.py                                format spec, publisher presets, guidelines parser, journal lookup
  template.py                            learn format from a Word template or sample PDF
  refs.py                                reference parsing/styling, in-text citation conversion
  pipeline.py checks.py                  numbering, captions, citations, compliance checks
  algo.py texblocks.py raw_import.py texchars.py      algorithm detection, LaTeX block protection, verbatim XML import
  render_docx.py render_latex.py         writers
  engine.py                              high-level API
tests/make_fixtures.py, tests/run_tests.py
```

## Notes

- With a Word template, the output is built inside the template itself, so its styles, fonts, spacing, columns,
  headers and footers carry over exactly; the template's sample text is discarded.
- PDF input recovers text, fonts, raster images and tables; display equations become exact images (not editable)
  and vector figures need re-inserting (both flagged in the report).
- MathType objects stay editable in the Word output; in the LaTeX export they are inserted as images.
- References that cannot be parsed confidently are kept as written and listed in the report.
- LaTeX output compiles on Overleaf with the default pdfLaTeX compiler: IEEEtran, elsarticle and llncs are used
  directly; Springer journals use `article` (Springer's sn-jnl class is not on Overleaf/TeX Live). Unicode symbols
  (Greek, ≤, ×, superscripts…) are converted to LaTeX commands automatically.
