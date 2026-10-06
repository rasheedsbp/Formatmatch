"""Unicode → pdfLaTeX-safe text (works with the default Overleaf compiler)."""
from __future__ import annotations

import re
import unicodedata

# maths symbols: LaTeX maths command (used inside $…$ in text, bare in maths mode)
MATH_SYM = {
    "±": r"\pm", "∓": r"\mp", "×": r"\times", "÷": r"\div", "·": r"\cdot", "⋅": r"\cdot", "∙": r"\cdot",
    "∘": r"\circ", "≤": r"\leq", "≥": r"\geq", "≦": r"\leqq", "≧": r"\geqq", "≠": r"\neq", "≈": r"\approx",
    "≃": r"\simeq", "≅": r"\cong", "≡": r"\equiv", "∼": r"\sim", "∝": r"\propto", "≪": r"\ll", "≫": r"\gg",
    "→": r"\rightarrow", "←": r"\leftarrow", "↔": r"\leftrightarrow", "⇒": r"\Rightarrow", "⇐": r"\Leftarrow",
    "⇔": r"\Leftrightarrow", "↑": r"\uparrow", "↓": r"\downarrow", "↦": r"\mapsto", "⟶": r"\longrightarrow",
    "∈": r"\in", "∉": r"\notin", "∋": r"\ni", "⊂": r"\subset", "⊃": r"\supset", "⊆": r"\subseteq",
    "⊇": r"\supseteq", "∩": r"\cap", "∪": r"\cup", "∅": r"\emptyset", "∀": r"\forall", "∃": r"\exists",
    "¬": r"\neg", "∧": r"\wedge", "∨": r"\vee", "⊕": r"\oplus", "⊗": r"\otimes", "⊙": r"\odot", "⊥": r"\perp",
    "∥": r"\parallel", "∠": r"\angle", "∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∬": r"\iint", "∮": r"\oint",
    "∂": r"\partial", "∇": r"\nabla", "√": r"\surd", "∞": r"\infty", "ℓ": r"\ell", "ℏ": r"\hbar",
    "ℝ": r"\mathbb{R}", "ℕ": r"\mathbb{N}", "ℤ": r"\mathbb{Z}", "ℚ": r"\mathbb{Q}", "ℂ": r"\mathbb{C}",
    "′": r"\prime", "″": r"\prime\prime", "−": "-", "∗": r"\ast", "⋆": r"\star", "†": r"\dagger", "‡": r"\ddagger",
    "∶": ":", "∣": r"\mid", "⌊": r"\lfloor", "⌋": r"\rfloor", "⌈": r"\lceil", "⌉": r"\rceil", "⟨": r"\langle",
    "⟩": r"\rangle", "‖": r"\|", "°": r"^\circ", "µ": r"\mu", "Å": r"\mathring{A}", "∴": r"\therefore",
    "∵": r"\because", "⊤": r"\top", "□": r"\square", "△": r"\triangle", "▷": r"\triangleright",
    "◁": r"\triangleleft", "⋯": r"\cdots", "⋮": r"\vdots", "⋱": r"\ddots", "…": r"\ldots", "ϵ": r"\epsilon",
    "ϑ": r"\vartheta", "ϕ": r"\phi", "ϖ": r"\varpi", "ϱ": r"\varrho", "ς": r"\varsigma",
}
GREEK = {
    "alpha": "alpha", "beta": "beta", "gamma": "gamma", "delta": "delta", "epsilon": "varepsilon", "zeta": "zeta",
    "eta": "eta", "theta": "theta", "iota": "iota", "kappa": "kappa", "lamda": "lambda", "lambda": "lambda",
    "mu": "mu", "nu": "nu", "xi": "xi", "omicron": "o", "pi": "pi", "rho": "rho", "sigma": "sigma", "tau": "tau",
    "upsilon": "upsilon", "phi": "varphi", "chi": "chi", "psi": "psi", "omega": "omega",
}
GREEK_CAP_LATIN = {"alpha": "A", "beta": "B", "epsilon": "E", "zeta": "Z", "eta": "H", "iota": "I", "kappa": "K",
                   "mu": "M", "nu": "N", "omicron": "O", "rho": "P", "tau": "T", "upsilon": r"\Upsilon", "chi": "X"}
SUP = {"⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
       "⁺": "+", "⁻": "-", "⁼": "=", "⁽": "(", "⁾": ")", "ⁿ": "n", "ⁱ": "i"}
SUB = {"₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
       "₊": "+", "₋": "-", "₌": "=", "₍": "(", "₎": ")", "ₐ": "a", "ₑ": "e", "ₒ": "o", "ₓ": "x", "ᵢ": "i", "ⱼ": "j",
       "ₖ": "k", "ₙ": "n", "ₘ": "m", "ₜ": "t"}
TEXT_SYM = {
    "“": "``", "”": "''", "„": ",,", "‘": "`", "’": "'", "‚": ",", "–": "--", "—": "---", "―": "---", "‐": "-",
    "‑": "-", "‒": "--", "…": r"\ldots{}", " ": "~", " ": r"\,", " ": r"\,", " ": " ",
    " ": r"\quad{}", "​": "", "‌": "", "‍": "", "⁠": "", "﻿": "", "­": "",
    "•": r"\textbullet{}", "§": r"\S{}", "¶": r"\P{}", "©": r"\textcopyright{}", "®": r"\textregistered{}",
    "™": r"\texttrademark{}", "€": r"\texteuro{}", "£": r"\pounds{}", "¥": r"\textyen{}", "₹": "Rs.",
    "¢": r"\textcent{}", "‰": r"\textperthousand{}", "℃": r"$^\circ$C", "℉": r"$^\circ$F", "№": "No.",
    "✓": r"$\checkmark$", "✔": r"$\checkmark$", "✗": r"$\times$", "✘": r"$\times$", "★": r"$\star$",
    "«": r"\guillemotleft{}", "»": r"\guillemotright{}", "¡": r"\textexclamdown{}", "¿": r"\textquestiondown{}",
    "ﬁ": "fi", "ﬂ": "fl", "ﬀ": "ff", "ﬃ": "ffi", "ﬄ": "ffl", " ": " ", " ": " ", "\t": " ",
}
SPECIAL = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#", "_": r"\_",
           "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}", "^": r"\textasciicircum{}", "<": r"\textless{}",
           ">": r"\textgreater{}", "|": r"\textbar{}"}


def _t1_ok(ch: str) -> bool:
    """Characters pdfLaTeX (utf8 + T1) typesets out of the box: Latin-1 and Latin Extended-A letters."""
    o = ord(ch)
    return 0xC0 <= o <= 0x17F and ch not in "×÷" or ch in "ß¿¡ªº"


def _greek(ch: str):
    try:
        name = unicodedata.name(ch)
    except ValueError:
        return None
    m = re.match(r"GREEK (SMALL|CAPITAL) LETTER (\w+)$", name)
    if not m:
        m2 = re.match(r"MATHEMATICAL (?:ITALIC|BOLD|BOLD ITALIC) (SMALL|CAPITAL) (\w+)$", name)
        if m2 and m2.group(2).lower() in GREEK:
            m = m2
        else:
            return None
    small, base = m.group(1) == "SMALL", m.group(2).lower()
    if base not in GREEK:
        return None
    if small:
        return "\\" + GREEK[base]
    if base in GREEK_CAP_LATIN:
        return GREEK_CAP_LATIN[base]
    cmd = GREEK[base].replace("var", "")
    return "\\" + cmd[0].upper() + cmd[1:]


def math_char(ch: str):
    """LaTeX maths-mode code for one non-ASCII character, or None."""
    if ch in MATH_SYM:
        return MATH_SYM[ch]
    g = _greek(ch)
    if g:
        return g
    if ch in SUP:
        return "^{" + SUP[ch] + "}"
    if ch in SUB:
        return "_{" + SUB[ch] + "}"
    try:
        name = unicodedata.name(ch)
    except ValueError:
        return None
    m = re.match(r"MATHEMATICAL (BOLD |ITALIC |BOLD ITALIC |SCRIPT |DOUBLE-STRUCK |FRAKTUR |SANS-SERIF )*(SMALL|CAPITAL) ([A-Z])$", name)
    if m:
        letter = m.group(3) if m.group(2) == "CAPITAL" else m.group(3).lower()
        style = (m.group(1) or "").strip()
        wrap = {"BOLD": r"\mathbf", "DOUBLE-STRUCK": r"\mathbb", "SCRIPT": r"\mathcal", "FRAKTUR": r"\mathfrak"}.get(style)
        return f"{wrap}{{{letter}}}" if wrap else letter
    return None


def tex_text(s: str, unknown: set = None) -> str:
    """Escape plain text for pdfLaTeX."""
    out = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch in SPECIAL:
            out.append(SPECIAL[ch])
        elif ord(ch) < 128:
            out.append(ch if ch >= " " or ch == "\n" else " ")
        elif ch in TEXT_SYM:
            out.append(TEXT_SYM[ch])
        elif ch in SUP or ch in SUB:
            table, cmd = (SUP, "textsuperscript") if ch in SUP else (SUB, "textsubscript")
            j = i
            buf = ""
            while j < n and s[j] in table:
                buf += table[s[j]]
                j += 1
            out.append(f"\\{cmd}{{{buf}}}")
            i = j
            continue
        elif _t1_ok(ch):
            out.append(ch)
        else:
            m = math_char(ch)
            if m:
                out.append(f"${m}$")
            else:
                d = unicodedata.normalize("NFKD", ch)
                base = "".join(c for c in d if not unicodedata.combining(c))
                if base and all(ord(c) < 128 or _t1_ok(c) for c in base) and base != ch:
                    out.append(tex_text(base))
                else:
                    if unknown is not None:
                        unknown.add(ch)
                    out.append("?")
        i += 1
    s2 = "".join(out)
    return s2.replace("$$", "")          # adjacent maths symbols merge into one maths group


def tex_math(s: str, unknown: set = None) -> str:
    """Make a LaTeX maths string (e.g. from pandoc) safe for pdfLaTeX: replace stray Unicode."""
    out = []
    for ch in s:
        if ord(ch) < 128:
            out.append(ch)
            continue
        m = math_char(ch)
        if m is None and ch in TEXT_SYM:
            m = {"–": "-", "—": "-", " ": "~", " ": r"\,", "​": ""}.get(ch, "")
        if m is None:
            if _t1_ok(ch):
                m = r"\text{" + ch + "}"
            else:
                if unknown is not None:
                    unknown.add(ch)
                m = "?"
        out.append(m + (" " if m and m[-1].isalpha() and m.startswith("\\") else ""))
    return "".join(out)
