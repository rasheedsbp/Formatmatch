"""Intermediate document model shared by all readers and writers."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, List, Optional

# Block roles
TITLE = "title"
AUTHOR = "author"
AFFIL = "affiliation"
ABSTRACT = "abstract"
KEYWORDS = "keywords"
HEADING = "heading"
PARA = "para"
LIST = "list"
TABLE = "table"
FIGURE = "figure"
CAPTION = "caption"
EQUATION = "equation"
REFERENCE = "reference"
OTHER_FRONT = "front"  # correspondence, dates, etc.
ABSTRACT_HEADING = "abstract_heading"  # standalone "Abstract" label line
ALGORITHM = "algorithm"


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    sup: bool = False
    sub: bool = False
    # inline equation: OMML XML string (docx source) and/or LaTeX
    omml: Optional[str] = None
    latex: Optional[str] = None
    # citation marker produced by the citation engine
    cite: bool = False
    # verbatim source run (e.g. MathType / Equation Editor OLE object) + its preview image (blob, ext)
    raw: Any = None
    img: Any = None

    @property
    def is_math(self) -> bool:
        return self.omml is not None or self.latex is not None or self.raw is not None or \
            (self.img is not None and not self.text)


@dataclass
class Block:
    role: str
    runs: List[Run] = field(default_factory=list)
    level: int = 0                       # heading level (1..4)
    ordered: bool = False                # list
    list_level: int = 0
    rows: Optional[List[List[str]]] = None   # table cells (plain text)
    image: Optional[bytes] = None        # figure
    image_ext: str = "png"
    width_in: Optional[float] = None
    height_in: Optional[float] = None
    cap_kind: str = ""                   # caption: 'figure' | 'table'
    cap_num: Optional[int] = None
    label: str = ""                      # e.g. abstract sub-heading "Purpose"
    style_name: str = ""                 # original style (diagnostics)
    eq_num: Optional[str] = None         # equation number text
    orig: str = ""                       # original paragraph text (before cleaning)
    src: Any = None                      # source XML element (docx) – used by template analysis
    src_size: float = 0.0
    src_font: str = ""
    src_bold: bool = False
    prefix: str = ""                     # generated numbering/label text (heading no., caption label)
    raw: Optional[list] = None           # verbatim source XML elements (equations, algorithms, complex tables)
    lines: Optional[list] = None         # algorithm lines: [(indent_level, [Run])]
    rows_runs: Optional[list] = None     # table cells as runs (keeps inline maths/formatting)
    tex_src: str = ""                    # original LaTeX source (algorithms from .tex input)

    @property
    def text(self) -> str:
        return "".join(r.text if not r.is_math else (r.latex or "[math]") for r in self.runs)

    def set_text(self, text: str):
        self.runs = [Run(text)]


@dataclass
class Document:
    blocks: List[Block] = field(default_factory=list)
    source_type: str = ""
    warnings: List[str] = field(default_factory=list)
    src_docx: Any = None                 # python-docx Document the blocks were read from (for raw XML parts)
    tex_preamble: List[str] = field(default_factory=list)  # \\usepackage lines needed by verbatim LaTeX blocks

    def by_role(self, role):
        return [b for b in self.blocks if b.role == role]

    def first(self, role):
        for b in self.blocks:
            if b.role == role:
                return b
        return None
