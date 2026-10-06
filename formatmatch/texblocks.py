"""LaTeX input: protect algorithm and display-equation environments from pandoc.

Each environment is replaced by a marker paragraph before conversion and restored afterwards as a block that
carries the *original LaTeX source* (used verbatim in the LaTeX output) plus a Word rendering (OMML maths via
pandoc, algorithm lines parsed from algorithmic / algorithm2e commands).
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from typing import Dict, List, Tuple

import docx

from . import model as M
from .model import Run

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
ALGO_ENVS = r"algorithm\*?|algorithm2e\*?|algorithmic|algorithmicx|procedure"
MATH_ENVS = r"equation\*?|align\*?|gather\*?|multline\*?|eqnarray\*?|flalign\*?|alignat\*?"
PKG_RE = re.compile(r"^\s*\\usepackage(?:\[[^\]]*\])?\{[^}]*(?:algorithm|algpseudocode|algorithmicx|algorithmic|"
                    r"algcompatible|algpascal|algc)[^}]*\}.*$", re.M)
MACRO_RE = re.compile(r"^\s*\\(?:algrenewcommand|algnewcommand|algdef|SetKw\w*|SetAlgo\w*|DontPrintSemicolon|"
                      r"SetAlFnt|renewcommand\{\\algorithmic\w*\}).*$", re.M)


def _strip_comments(s: str) -> str:
    return re.sub(r"(?<!\\)%.*", "", s)


def expand_inputs(tex: str, root: str, depth: int = 0) -> str:
    """Inline \\input{…} / \\include{…} files so environments in sub-files are protected too."""
    if depth > 5:
        return tex

    def sub(m):
        name = m.group(2).strip()
        for cand in (name, name + ".tex"):
            path = os.path.join(root, cand)
            if os.path.isfile(path):
                try:
                    return expand_inputs(open(path, encoding="utf-8", errors="ignore").read(), root, depth + 1)
                except OSError:
                    break
        return m.group(0)
    return re.sub(r"\\(input|include)\{([^}]+)\}", sub, tex)


def protect(tex: str):
    """Return (modified_tex, blocks, preamble_lines). blocks[k] = dict(kind, src, number)."""
    blocks: List[dict] = []
    head, sep, body = tex.partition("\\begin{document}")
    if not sep:
        head, body = "", tex
    preamble = PKG_RE.findall(head) + MACRO_RE.findall(head)
    preamble = [p.strip() for p in preamble]

    # 1. algorithms (outermost environments only)
    def algo_sub(m):
        k = len(blocks)
        blocks.append({"kind": "algo", "src": m.group(0)})
        return f"\n\nFMBLOCK{k}X\n\n"
    body = re.sub(r"\\begin\{(" + ALGO_ENVS + r")\}.*?\\end\{\1\}", algo_sub, body, flags=re.S)

    # 2. display maths
    eq_counter = [0]

    def numbers_for(env: str, src: str) -> str:
        if env.endswith("*") or env in ("displaymath",):
            return ""
        rows = re.split(r"\\\\(?!\w)", src) if env not in ("equation", "multline") else [src]
        start = eq_counter[0] + 1
        for r in rows:
            if r.strip() and not re.search(r"\\(?:nonumber|notag)\b", r):
                eq_counter[0] += 1
        end = eq_counter[0]
        if end < start:
            return ""
        return str(start) if end == start else f"{start}–{end}"

    labels: Dict[str, str] = {}

    def math_sub(m):
        env = m.group(1) or ""
        src = m.group(0)
        num = numbers_for(env, m.group(2) or "") if env else ""
        for lab in re.findall(r"\\label\{([^}]*)\}", src):
            labels[lab] = num.split("–")[0] if num else ""
        k = len(blocks)
        blocks.append({"kind": "eq", "src": src, "env": env, "number": num, "body": m.group(2) if env else m.group(3)})
        return f"\n\nFMBLOCK{k}X\n\n"
    body = re.sub(r"\\begin\{(" + MATH_ENVS + r")\}(.*?)\\end\{\1\}|\\\[(?P<x>)(.*?)\\\]",
                  lambda m: math_sub(_EqMatch(m)), body, flags=re.S)

    # algorithm labels -> numbers (in order of appearance)
    n_alg = 0
    for b in blocks:
        if b["kind"] == "algo" and re.search(r"\\caption", b["src"]):
            n_alg += 1
            b["number"] = str(n_alg)
            for lab in re.findall(r"\\label\{([^}]*)\}", b["src"]):
                labels[lab] = str(n_alg)

    # \ref / \eqref to protected blocks become literal numbers (pandoc cannot resolve them)
    def ref_sub(m):
        lab = m.group(2)
        if lab not in labels:
            return m.group(0)
        n = labels[lab]
        return f"({n})" if m.group(1) == "eqref" else n
    body = re.sub(r"\\(eqref|ref|autoref|cref|Cref)\{([^}]*)\}", ref_sub, body)
    return head + sep + body, blocks, preamble


class _EqMatch:
    """Normalise the two alternatives of the display-maths regex."""
    def __init__(self, m):
        self._m = m
        if m.group(1):
            self.g = (m.group(0), m.group(1), m.group(2), None)
        else:
            self.g = (m.group(0), None, None, m.group(4))

    def group(self, i):
        return self.g[i]


# ---------------------------------------------------------------- LaTeX maths → OMML (batched)
def latex_to_omml(items: List[str]) -> List[str]:
    """Convert LaTeX maths snippets to OMML XML strings via pandoc ("" when conversion fails)."""
    if not items or not shutil.which("pandoc"):
        return [""] * len(items)
    md = []
    for k, t in enumerate(items):
        md.append(f"FMEQ{k}X\n\n$${t}$$\n")
    out = [""] * len(items)
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "m.md"); dst = os.path.join(td, "m.docx")
        open(src, "w", encoding="utf-8").write("\n".join(md))
        r = subprocess.run(["pandoc", src, "-f", "markdown+tex_math_dollars", "-o", dst], capture_output=True,
                           timeout=180)
        if r.returncode != 0:
            return out
        d = docx.Document(dst)
        cur = None
        from lxml import etree
        for p in d.paragraphs:
            m = re.fullmatch(r"FMEQ(\d+)X", p.text.strip())
            if m:
                cur = int(m.group(1)); continue
            if cur is not None:
                om = next(p._p.iter(f"{{{NS_M}}}oMathPara"), None)
                if om is None:
                    om = next(p._p.iter(f"{{{NS_M}}}oMath"), None)
                if om is not None and not out[cur]:
                    out[cur] = etree.tostring(om, encoding="unicode")
    return out


def _env_math_for_pandoc(b: dict) -> str:
    env = (b.get("env") or "").rstrip("*")
    body = b.get("body") or ""
    body = re.sub(r"\\label\{[^}]*\}|\\(?:nonumber|notag)\b", "", body)
    if env in ("align", "flalign", "alignat", "eqnarray"):
        body = re.sub(r"^\{\d+\}", "", body) if env == "alignat" else body
        body = body.replace("&=&", "&=")
        return "\\begin{aligned}" + body + "\\end{aligned}"
    if env in ("gather", "multline"):
        return "\\begin{gathered}" + body + "\\end{gathered}"
    return body


# ---------------------------------------------------------------- algorithmic → lines
CMD_TEXT = {
    "Require": "Input:", "Ensure": "Output:", "Input": "Input:", "Output": "Output:", "KwIn": "Input:",
    "KwOut": "Output:", "KwData": "Data:", "KwResult": "Result:", "REQUIRE": "Input:", "ENSURE": "Output:",
}


def _brace_arg(s: str, i: int) -> Tuple[str, int]:
    """Read a {…} group starting at s[i] == '{'. Returns (content, index after group)."""
    if i >= len(s) or s[i] != "{":
        return "", i
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{" and (j == 0 or s[j - 1] != "\\"):
            depth += 1
        elif s[j] == "}" and s[j - 1] != "\\":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
    return s[i + 1:], len(s)


def parse_algorithm(src: str):
    """(caption, [(indent, text_with_$math$)], numbered)"""
    s = _strip_comments(src)
    cap = ""
    m = re.search(r"\\caption\{", s)
    if m:
        cap, _ = _brace_arg(s, m.end() - 1)
    numbered = bool(re.search(r"\\begin\{algorithmic\}\[\s*1\s*\]", s)) or "\\LinesNumbered" in s
    body = re.sub(r"\\caption\{", "", s, count=0)
    lines: List[Tuple[int, str]] = []
    indent = 0
    is_a2e = "\\SetKw" in s or "\\KwIn" in s or "\\eIf" in s or "algorithm2e" in s or "\\;" in s
    if is_a2e:
        return cap, _parse_a2e(s), numbered
    tokens = re.split(r"(\\(?:State|STATE|Statex|Require|Ensure|REQUIRE|ENSURE|Input|Output|If|IF|ElsIf|ELSIF|"
                      r"Else|ELSE|EndIf|ENDIF|For|FOR|ForAll|FORALL|EndFor|ENDFOR|While|WHILE|EndWhile|ENDWHILE|"
                      r"Repeat|REPEAT|Until|UNTIL|Loop|LOOP|EndLoop|ENDLOOP|Function|EndFunction|Procedure|"
                      r"EndProcedure|Return|RETURN|Comment|COMMENT|Call|label|caption|begin|end)\b\*?)", s)
    i = 0
    cur_text = None

    def flush():
        nonlocal cur_text
        if cur_text is not None and cur_text.strip():
            lines.append((max(0, indent_at[0]), cur_text.strip()))
        cur_text = None
    indent_at = [0]
    k = 0
    while k < len(tokens):
        tok = tokens[k]
        rest = tokens[k + 1] if k + 1 < len(tokens) else ""
        if not tok.startswith("\\"):
            if cur_text is not None:
                cur_text += tok
            k += 1
            continue
        name = tok[1:].rstrip("*")
        args = []
        j = 0
        r = rest
        while r[j:j + 1] == "{" or (r[j:j + 1].isspace() and r[j:].lstrip().startswith("{") and len(args) < 2 and
                                     name in ("If", "IF", "ElsIf", "ELSIF", "For", "FOR", "ForAll", "FORALL", "While",
                                              "WHILE", "Until", "UNTIL", "Function", "Procedure", "Comment",
                                              "COMMENT", "Call", "label", "caption", "begin", "end")):
            j += len(r[j:]) - len(r[j:].lstrip())
            a, j = _brace_arg(r, j)
            args.append(a)
            if name not in ("Function", "Procedure", "Call"):
                break
        tail = r[j:]
        tokens[k + 1] = tail
        if name in ("label", "caption", "begin", "end"):
            k += 1
            continue
        if name in ("Comment", "COMMENT"):
            if cur_text is not None:
                cur_text += "  ▷ " + (args[0] if args else "")
            k += 1
            continue
        flush()
        lower = name.lower()
        if lower in ("endif", "endfor", "endwhile", "endloop", "endfunction", "endprocedure"):
            indent = max(0, indent - 1)
            indent_at[0] = indent
            word = {"endif": "end if", "endfor": "end for", "endwhile": "end while", "endloop": "end loop",
                    "endfunction": "end function", "endprocedure": "end procedure"}[lower]
            cur_text = word
            flush()
        elif lower in ("if", "for", "forall", "while", "loop", "repeat", "function", "procedure"):
            indent_at[0] = indent
            a = args[0] if args else ""
            if lower == "if":
                cur_text = f"if {a} then"
            elif lower in ("for", "forall"):
                cur_text = f"for {'all ' if lower == 'forall' else ''}{a} do"
            elif lower == "while":
                cur_text = f"while {a} do"
            elif lower in ("function", "procedure"):
                cur_text = f"{lower} {a}({args[1] if len(args) > 1 else ''})"
            else:
                cur_text = lower
            indent += 1
        elif lower in ("elsif", "else"):
            indent_at[0] = max(0, indent - 1)
            cur_text = f"else if {args[0]} then" if lower == "elsif" and args else "else"
        elif lower == "until":
            indent = max(0, indent - 1)
            indent_at[0] = indent
            cur_text = f"until {args[0] if args else ''}"
        elif lower == "return":
            indent_at[0] = indent
            cur_text = "return"
        elif name in CMD_TEXT:
            indent_at[0] = 0
            cur_text = CMD_TEXT[name] + " "
        else:   # State / Statex / Call
            indent_at[0] = indent
            cur_text = ""
        k += 1
    flush()
    return cap, [(lvl, re.sub(r"\s+", " ", t)) for lvl, t in lines], numbered


def _parse_a2e(s: str):
    """algorithm2e: statements end with \\; and blocks are {…} arguments."""
    s = re.sub(r"\\(?:caption|label)\{[^}]*\}", "", s)
    s = re.sub(r"\\(?:begin|end)\{algorithm2e\*?\}(\[[^\]]*\])?", "", s)
    s = re.sub(r"\\(?:SetKw\w*|DontPrintSemicolon|SetAlgo\w*|LinesNumbered|SetAlFnt)(\{[^}]*\})*", "", s)
    out: List[Tuple[int, str]] = []

    def walk(text: str, ind: int):
        i = 0
        buf = ""
        while i < len(text):
            m = re.match(r"\\(KwIn|KwOut|KwData|KwResult|Input|Output)\{", text[i:])
            if m:
                a, j = _brace_arg(text, i + m.end() - 1)
                out.append((0, CMD_TEXT.get(m.group(1), m.group(1) + ":") + " " + a))
                i = j
                continue
            m = re.match(r"\\(eIf|If|uIf|ElseIf|uElseIf|Else|uElse|For|ForEach|ForAll|While|Repeat|Fn|Begin)\b\s*", text[i:])
            if m:
                if buf.strip():
                    out.append((ind, buf.strip())); buf = ""
                name = m.group(1)
                j = i + m.end()
                args = []
                while j < len(text) and text[j] == "{":
                    a, j = _brace_arg(text, j)
                    args.append(a)
                    while j < len(text) and text[j].isspace():
                        j += 1
                    if len(args) >= (3 if name == "eIf" else 2 if name not in ("Else", "uElse", "Begin") else 1):
                        break
                if name in ("eIf", "If", "uIf"):
                    out.append((ind, f"if {args[0] if args else ''} then"))
                    if len(args) > 1:
                        walk(args[1], ind + 1)
                    if name == "eIf" and len(args) > 2:
                        out.append((ind, "else")); walk(args[2], ind + 1)
                    out.append((ind, "end if"))
                elif name in ("ElseIf", "uElseIf"):
                    out.append((ind, f"else if {args[0] if args else ''} then"))
                    if len(args) > 1:
                        walk(args[1], ind + 1)
                elif name in ("Else", "uElse"):
                    out.append((ind, "else"))
                    if args:
                        walk(args[0], ind + 1)
                elif name in ("For", "ForEach", "ForAll", "While"):
                    word = {"For": "for", "ForEach": "foreach", "ForAll": "for all", "While": "while"}[name]
                    out.append((ind, f"{word} {args[0] if args else ''} do"))
                    if len(args) > 1:
                        walk(args[1], ind + 1)
                    out.append((ind, f"end {word.split()[0]}"))
                elif name == "Repeat":
                    out.append((ind, "repeat"))
                    if len(args) > 1:
                        walk(args[1], ind + 1)
                    out.append((ind, f"until {args[0] if args else ''}"))
                else:
                    if args:
                        walk(args[-1], ind + 1 if name != "Begin" else ind)
                i = j
                continue
            if text.startswith("\\;", i):
                if buf.strip():
                    out.append((ind, buf.strip()))
                buf = ""
                i += 2
                continue
            m = re.match(r"\\(Return|KwRet)\{", text[i:])
            if m:
                a, j = _brace_arg(text, i + m.end() - 1)
                buf += "return " + a
                i = j
                continue
            m = re.match(r"\\(?:tcc|tcp)\*?\{", text[i:])
            if m:
                a, j = _brace_arg(text, i + m.end() - 1)
                buf += "  ▷ " + a
                i = j
                continue
            buf += text[i]
            i += 1
        if buf.strip():
            out.append((ind, buf.strip()))
    walk(s, 0)
    return [(l, re.sub(r"\\KwTo\b", "to", re.sub(r"\s+", " ", t))) for l, t in out if t.strip()]


# ---------------------------------------------------------------- text with $math$ → runs
_TEXT_CMDS = [(r"\\textbf\{([^{}]*)\}", "b"), (r"\\textit\{([^{}]*)\}", "i"), (r"\\emph\{([^{}]*)\}", "i"),
              (r"\\texttt\{([^{}]*)\}", ""), (r"\\text\{([^{}]*)\}", "")]


def _plain(t: str) -> str:
    t = re.sub(r"\\(?:KwTo)\b", "to", t)
    t = re.sub(r"\\(?:gets|leftarrow)\b", "←", t)
    t = t.replace("\\&", "&").replace("\\%", "%").replace("\\_", "_").replace("~", " ")
    t = re.sub(r"\\[a-zA-Z]+\*?", "", t)
    return t.replace("{", "").replace("}", "")


def text_to_runs(t: str, math_pool: list) -> List[Run]:
    """Split "for $i=1$ to $n$" into text runs and maths runs (maths collected in math_pool for OMML)."""
    runs: List[Run] = []
    pos = 0
    for m in re.finditer(r"(?<!\\)\$(.+?)(?<!\\)\$|\\\((.+?)\\\)", t):
        if m.start() > pos:
            runs += _styled(t[pos:m.start()])
        lx = (m.group(1) or m.group(2)).strip()
        r = Run(latex=lx)
        math_pool.append(r)
        runs.append(r)
        pos = m.end()
    if pos < len(t):
        runs += _styled(t[pos:])
    return [r for r in runs if r.text or r.is_math]


def _styled(t: str) -> List[Run]:
    out = []
    pos = 0
    pat = re.compile(r"\\(textbf|textit|emph|texttt|text)\{([^{}]*)\}")
    for m in pat.finditer(t):
        if m.start() > pos:
            out.append(Run(_plain(t[pos:m.start()])))
        out.append(Run(_plain(m.group(2)), bold=m.group(1) == "textbf", italic=m.group(1) in ("textit", "emph")))
        pos = m.end()
    if pos < len(t):
        out.append(Run(_plain(t[pos:])))
    return out


def restore(doc: M.Document, blocks: List[dict], preamble: List[str]):
    """Replace marker paragraphs with ALGORITHM / EQUATION blocks."""
    if not blocks:
        return
    math_pool: List[Run] = []
    built = {}
    eq_items = []
    for k, b in enumerate(blocks):
        if b["kind"] == "eq":
            eq_items.append((k, _env_math_for_pandoc(b)))
            continue
        cap, lines, numbered = parse_algorithm(b["src"])
        cap_runs = text_to_runs(cap, math_pool) if cap else []
        title = f"Algorithm {b.get('number', '')}".strip()
        head_runs = [Run(title + (": " if cap else ""), bold=True)] + cap_runs
        line_runs = [(lvl, text_to_runs(t, math_pool)) for lvl, t in lines]
        if numbered:
            n, numbered_runs = 0, []
            for (lvl, rs), (_, t) in zip(line_runs, lines):
                if re.match(r"^(Input|Output|Data|Result):", t):
                    numbered_runs.append((lvl, rs))
                else:
                    n += 1
                    numbered_runs.append((lvl, [Run(f"{n}: ")] + rs))
            line_runs = numbered_runs
        built[k] = M.Block(M.ALGORITHM, runs=head_runs, lines=line_runs, tex_src=b["src"],
                           orig=title + (": " + _plain(cap) if cap else ""), ordered=numbered)
    # OMML for every maths snippet in one pandoc call
    snippets = [x for _, x in eq_items] + [r.latex for r in math_pool]
    omml = latex_to_omml(snippets)
    for (k, _), om in zip(eq_items, omml[:len(eq_items)]):
        b = blocks[k]
        run = Run(omml=om) if om else Run(latex=b.get("body") or "")
        built[k] = M.Block(M.EQUATION, runs=[run], eq_num=b.get("number") or None, tex_src=b["src"])
    for r, om in zip(math_pool, omml[len(eq_items):]):
        if om:
            r.omml = om
    new_blocks = []
    for blk in doc.blocks:
        m = re.fullmatch(r"\s*FMBLOCK(\d+)X\s*", blk.text or "")
        if m and int(m.group(1)) in built:
            new_blocks.append(built[int(m.group(1))])
            continue
        if m:
            continue
        new_blocks.append(blk)
    doc.blocks = new_blocks
    doc.tex_preamble = list(dict.fromkeys(preamble))
