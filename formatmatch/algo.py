"""Detection helpers for algorithm / pseudocode blocks."""
from __future__ import annotations

import re

ALGO_HEAD_RE = re.compile(
    r"^\s*(?:algorithm|pseudo-?\s?code|procedure)\s*(?:[\dIVX]+(?:\.\d+)?)?\s*(?P<sep>[:.\-–—|])?\s*(?P<rest>.*)$",
    re.I | re.S,
)
KEYWORDS = (
    r"input|inputs|output|outputs|require|requires|ensure|ensures|begin|end|endfor|endwhile|endif|end\s+for|"
    r"end\s+while|end\s+if|end\s+procedure|end\s+function|initiali[sz]e|initiali[sz]ation|for|foreach|for\s+each|"
    r"while|if|else|elif|else\s*if|then|do|return|repeat|until|loop|break|continue|set|let|compute|calculate|"
    r"update|select|evaluate|generate|sort|function|procedure|parameters?|data|result|stop|terminate|"
    r"termination|output:|call|apply|assign|obtain|construct|train|predict|define|declare|goto|step|phase|"
    r"next|print|exit|yield|switch|case|otherwise|endfunction|endprocedure|in\s+parallel"
)
KW_RE = re.compile(r"^\s*(?:" + KEYWORDS + r")\b", re.I)
# step verbs that start pseudo-code lines but (almost) never a section heading
STEP_RE = re.compile(r"^\s*(?:initiali[sz]e|compute|calculate|update|return|repeat|until|for|foreach|while|if|else|"
                     r"end|endfor|endwhile|endif|set|let|assign|generate|evaluate|select|sort|repeat|stop|"
                     r"terminate|input|output|begin|do|goto|print|obtain|apply|call|choose|randomly)\b", re.I)
LINENO_RE = re.compile(r"^\s*(?:\d{1,3}\s*[:.)]|\(\d{1,3}\)|step\s*\d+\s*[:.)-]?)\s*", re.I)
CODEISH_RE = re.compile(r"←|<-|:=|\+\+|==|!=|≤|≥|\b(?:do|then|end)\s*$|;\s*$|\bto\b.*\bdo\b", re.I)
STYLE_RE = re.compile(r"algorithm|pseudo|code|program|source|listing", re.I)
MONO_RE = re.compile(r"courier|consolas|mono|menlo|lucida console|inconsolata|source code", re.I)


def is_algo_head(text: str, bold: bool = False) -> bool:
    t = text.strip()
    if not t or len(t) > 250:
        return False
    m = ALGO_HEAD_RE.match(t)
    if not m:
        return False
    word = t.split()[0].lower()
    has_num = bool(re.match(r"^\s*\w[\w-]*\s*[\dIVX]+", t))
    if m.group("sep"):
        return has_num or word.startswith("pseudo") or bold
    # "Algorithm 1 Particle swarm optimisation" (no separator): needs a number and must not read as a sentence
    if has_num and (bold or len(t) < 110) and not re.search(r"\b(?:is|are|shows?|presents?|describes?|summari[sz]es?|gives?)\b", t, re.I):
        return True
    return False


def is_algo_line(text: str, indented: bool = False, numbered: bool = False, has_math: bool = False,
                 style: str = "", mono: bool = False) -> bool:
    t = text.strip()
    if style and STYLE_RE.search(style):
        return True
    if mono:
        return True
    if not t:
        return has_math
    if len(t) > 400:
        return False
    if LINENO_RE.match(t):
        rest = LINENO_RE.sub("", t, count=1)
        return len(rest) < 300
    if KW_RE.match(t) and len(t) < 160:
        return True
    if CODEISH_RE.search(t) and len(t) < 200:
        return True
    if (indented or numbered) and len(t) < 200:
        return True
    if has_math and len(t) < 160:
        return True
    return False


def strip_line_number(text: str):
    m = LINENO_RE.match(text)
    return (text[m.end():], True) if m else (text, False)
