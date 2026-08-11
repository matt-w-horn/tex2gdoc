#!/usr/bin/env python3
"""Convert a LaTeX paper to a .docx that opens cleanly in Word or Google Docs.

Pandoc does not run a TeX engine, so TikZ and PGFPlots figures never survive a
direct conversion. This script renders each such figure with a real TeX run,
substitutes the images into a working copy of the source, works around four
pandoc defects, sets the handful of styles that carry layout rather than
decoration, and then verifies the .docx that actually ships.

The four defects, each found by comparing an output against its source:

  * A table whose cell contains a nested `tabular` is dropped entirely — exit
    status 0, empty log, and an empty table shell left behind so that counting
    table elements still passes.
  * The caption of a starred `table*` or `figure*` is dropped, while the same
    caption in an unstarred float survives.
  * An `alltt` listing is read as a paragraph, so every run of spaces collapses
    and `structure Loop where` arrives as `structureLoopwhere`.
  * `\\ref` to a table, figure, or theorem is left as the raw label key, since
    pandoc numbers sections but not floats.

A figure is recompiled outside the paper, so whatever it reads from its
surroundings has to be measured first and carried across: \\columnwidth, and the
whole font size ladder. Both were wrong until 2026-08-09 and both were silent.
See `measure_page_metrics`.

Styling is deliberately minimal. Only styles that carry layout are touched, and
pandoc's own Heading1-3, Title and Caption definitions are left alone so Google
Docs maps them onto its native styles and its style menu keeps working.

Every OOXML edit and every OOXML check operates on a parsed tree, never on
serialized text. That is what lets the output survive a pandoc upgrade changing
how it spaces its tags, and it means a malformed document is caught by the
parser rather than quietly matching nothing. `ooxml` owns the one parser, the
element builder, and the tag and attribute names; nothing here constructs XML
from a string, and ruff blocks importing another parser.

The verification step is the point, and it runs against the output file rather
than a copy or an intermediate. All 21 checks are calibrated: `--self-test`
builds input designed to break each one and asserts it reports FAIL, and adding
a check without a calibration case fails that gate. It needs no pandoc, no TeX.

Usage:
    tex2gdoc main.tex
    tex2gdoc main.tex -o review.docx --bib references.bib
    tex2gdoc --self-test

Requires: pandoc, a TeX engine, and pdftocairo (Poppler).
    brew install pandoc poppler tectonic

tectonic is preferred over pdflatex: it is the XeTeX-based engine the target
paper is actually built with, so a measured \\columnwidth matches the real
document, and it fetches the packages a figure needs rather than requiring a
full TeX distribution.

Verified 2026-08-09 against a two-column, twenty-two page paper carrying every
construct above: pgfplots figures, a nested tabular inside a cell, code
listings, several hundred equations and dozens of citations. `examples/` holds
a small paper with the same shapes, which is the one to read first.

Math converts to native Word equations, and Google Docs renders them. Confirmed
2026-08-09 by generating a .docx, opening it in Docs, and reading it: inline
math, superscripts, Greek, set membership, and a summation with its limits above
and below all arrive correctly.

Two things do not survive, both worth knowing before sending a paper out:

  * `\\begin{cases}` is flattened. A two-branch piecewise definition arrives as a
    single run, with the branches and their conditions run together and no
    alignment. The target paper has one such block.
  * `\\mathbb{R}` loses its blackboard bold and arrives as an italic R. This does
    not apply inside an `alltt` listing, where the escapes are turned into
    literal characters before pandoc sees them.

Neither justifies rendering all math to images for the sake of two spans. If
that is ever wanted, it needs a Lua filter replacing Math nodes with rendered
images, because pandoc has no built-in option for it.

LibreOffice renders no OMML at all, not even `$x$`, so it shows blanks where the
math is and cannot be used to preview any of this.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import TYPE_CHECKING, Final

from .ooxml import (
    XML_SPACE_PRESERVE,
    Attr,
    Attribute,
    SafeElement,
    Tag,
    Value,
    attribute,
    child,
    children,
    descendants,
    make,
    parse_part,
    qn,
    serialize_part,
    value_of,
)

if TYPE_CHECKING:  # pragma: no cover - annotations only
    # Imported lazily at runtime inside convert(). `verification` reads this
    # module's constants, so a module-level import here would be a cycle.
    from .verification import Check

# --------------------------------------------------------------------------
# OOXML parsing
# --------------------------------------------------------------------------

# The namespace table, the hardened parser and the element builder all live in
# `ooxml`, which is the only module allowed to turn bytes into a tree.


class StyleId(StrEnum):
    """The pandoc style ids this converter is allowed to touch.

    Everything not named here keeps pandoc's own definition on purpose, so Word
    and Google Docs map Heading1-3, Title and Caption onto their native styles.
    """

    NORMAL = "Normal"
    BODY_TEXT = "BodyText"
    COMPACT = "Compact"
    CAPTIONED_FIGURE = "CaptionedFigure"
    FIGURE = "Figure"
    IMAGE_CAPTION = "ImageCaption"
    VERBATIM_CHAR = "VerbatimChar"
    SOURCE_CODE = "SourceCode"
    FIRST_PARAGRAPH = "FirstParagraph"
    BIBLIOGRAPHY = "Bibliography"
    ABSTRACT = "Abstract"
    BLOCK_TEXT = "BlockText"
    AUTHOR = "Author"
    SUBTITLE = "Subtitle"
    HEADING_1 = "Heading1"


# Keyword attribute values live in `ooxml.Value`; this one is a measurement.
HEADER_FOOTER_MARGIN = 360


def serialize_ooxml(root: SafeElement) -> str:
    return serialize_part(root).decode("utf-8")


def parse_ooxml(raw: str, part: str) -> SafeElement:
    """Parse an OOXML part through the package's one hardened parser.

    The hardening, and the reasoning behind it, live in `ooxml`. This wrapper
    stays because callers hold parts as `str`, and because `verify()` reads
    whatever path it is given, including a file that came back from a co-author,
    so provenance is not assumed anywhere in this package.
    """
    return parse_part(raw.encode("utf-8"), part)


# --------------------------------------------------------------------------
# Source parsing
# --------------------------------------------------------------------------

FIGURE_RE = re.compile(r"\\begin\{figure\*?\}.*?\\end\{figure\*?\}", re.DOTALL)
TIKZ_RE = re.compile(r"\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}", re.DOTALL)
LABEL_RE = re.compile(r"\\label\{([^}]+)\}")
CITE_RE = re.compile(r"\\cite[tp]?\*?(?:\[[^\]]*\])*\{([^}]+)\}")
INLINE_MATH_RE = re.compile(r"(?<!\\)\$(?!\$).+?(?<!\\)\$", re.DOTALL)
SECTION_RE = re.compile(r"^\s*\\section\*?\{", re.MULTILINE)

# Preamble lines worth carrying into a standalone figure. Anything else in the
# preamble belongs to the document class, not to the picture.
PREAMBLE_KEEP_RE = re.compile(
    r"^\s*\\("
    r"usepackage|usetikzlibrary|pgfplotsset|definecolor|newcommand|renewcommand"
    r"|providecommand|DeclareMathOperator|newlength|setlength|def"
    r")\b"
)

# Packages that belong to the paper but break or pollute a standalone figure.
PREAMBLE_DROP_PACKAGES = {"eso-pic", "draftwatermark", "geometry", "hyperref", "needspace"}


@dataclass
class Figure:
    """One figure environment in the source whose body is a TikZ picture."""

    index: int
    label: str
    block: str
    tikz: str
    png: Path | None = None
    fallback_preamble: bool = False


@dataclass
class SourceFacts:
    """What the source claims, for the output to be checked against."""

    tikz_figures: int = 0
    graphics_figures: int = 0
    top_level_tabulars: int = 0
    nested_tabulars: int = 0
    table_probes: list[str] = field(default_factory=list)
    caption_probes: list[str] = field(default_factory=list)
    listing_probes: list[str] = field(default_factory=list)
    citation_keys: set[str] = field(default_factory=set)
    inline_math: int = 0
    sections: int = 0
    body_chars: int = 0


def strip_comments(src: str) -> str:
    """Remove LaTeX line comments, preserving escaped percent signs."""
    out = []
    for line in src.split("\n"):
        pos, cut = 0, None
        while pos < len(line):
            i = line.find("%", pos)
            if i == -1:
                break
            if i > 0 and line[i - 1] == "\\":
                pos = i + 1
                continue
            cut = i
            break
        out.append(line if cut is None else line[:cut])
    return "\n".join(out)


def extract_braced(src: str, open_index: int) -> str:
    """Return the contents of the group whose opening brace is at open_index."""
    depth = 0
    for i in range(open_index, len(src)):
        if src[i] == "{":
            depth += 1
        elif src[i] == "}":
            depth -= 1
            if depth == 0:
                return src[open_index + 1 : i]
    return src[open_index + 1 :]


def find_tabular_spans(src: str) -> list[tuple[int, int, int]]:
    """Return (start, end, depth) for every tabular environment.

    Depth 0 means top level. Pandoc emits one Word table per top-level tabular,
    which is what makes this count checkable against the output.
    """
    token_re = re.compile(r"\\(begin|end)\{tabular\}")
    stack: list[int] = []
    spans: list[tuple[int, int, int]] = []
    for m in token_re.finditer(src):
        if m.group(1) == "begin":
            stack.append(m.start())
        elif stack:
            start = stack.pop()
            spans.append((start, m.end(), len(stack)))
    return spans


def table_probe(tabular: str, full_source: str) -> str | None:
    """Pick a distinctive word from a table that must survive into the output.

    Counting table elements is not enough: when pandoc drops a table's contents
    it can still leave an empty table shell behind, so the count check passes on
    an empty table. A word from inside the table is the thing that cannot be
    faked by an empty shell. A single token is used rather than a phrase because
    spacing between output runs is not stable.
    """
    text = re.sub(r"\\begin\{tabular\}(\[[^\]]*\])?\{[^}]*\}", " ", tabular)
    text = text.replace(r"\end{tabular}", " ").replace(r"\_", "_")
    # Keys are replaced by rendered text in the output, so they are never valid
    # probes: drop these macros together with their arguments.
    text = re.sub(r"\\(cite[tp]?\*?|label|ref|autoref|eqref)(\[[^\]]*\])*\{[^}]*\}", " ", text)
    text = re.sub(r"\\[a-zA-Z]+\s*", " ", text)  # other macro names, keeping their arguments
    text = re.sub(r"[{}&$\\]", " ", text)
    candidates = re.findall(r"[A-Za-z][A-Za-z0-9_]{7,}", text)
    if not candidates:
        return None
    # Prefer a word that occurs once in the whole source: a hit on it in the
    # output cannot have come from somewhere else.
    unique = [c for c in candidates if full_source.count(c) == 1]
    pool = unique or candidates
    return str(max(pool, key=len))


def collect_source_facts(src: str) -> SourceFacts:
    """Measure the source once, so the output has something to be checked against."""
    clean = strip_comments(src)
    body = clean.split(r"\begin{document}", 1)[-1]

    facts = SourceFacts()
    for block in FIGURE_RE.findall(body):
        if TIKZ_RE.search(block):
            facts.tikz_figures += 1
        elif r"\includegraphics" in block:
            facts.graphics_figures += 1

    for start, end, depth in find_tabular_spans(body):
        if depth == 0:
            facts.top_level_tabulars += 1
            probe = table_probe(body[start:end], clean)
            if probe:
                facts.table_probes.append(probe)
        else:
            facts.nested_tabulars += 1

    for group in CITE_RE.findall(body):
        for key in group.split(","):
            key = key.strip()
            if key:
                facts.citation_keys.add(key)

    facts.inline_math = len(INLINE_MATH_RE.findall(body))
    facts.sections = len(SECTION_RE.findall(body))

    for m in re.finditer(r"\\caption\{", body):
        cap = extract_braced(body, m.end() - 1)
        cap = re.sub(r"\\(label|ref|cite[tp]?)\*?\{[^}]*\}", " ", cap)
        cap = re.sub(r"\\[a-zA-Z]+|[{}$\\]", " ", cap)
        cap = re.sub(r"\s+", " ", cap).strip()
        # A contiguous phrase, not a bag of words: dropping a short word such as
        # "a" would join two words the output never puts side by side.
        phrase = re.search(r"[A-Za-z][A-Za-z ,'\-]{19,55}", cap)
        if phrase:
            facts.caption_probes.append(phrase.group(0).strip())

    # A listing check needs a contiguous phrase, not a bag of words: the defect
    # is collapsed spacing between tokens, and joining non-adjacent words would
    # report a failure the document does not have.
    for m in re.finditer(
        r"\\begin\{(alltt|verbatim|lstlisting)\}(.*?)\\end\{\1\}", body, re.DOTALL
    ):
        for line in m.group(2).split("\n"):
            plain = re.sub(r"\\\(.*?\\\)", "\x00", line)
            phrase = re.search(r"[A-Za-z][A-Za-z0-9_. ]{11,40}", plain)
            if phrase and " " in phrase.group(0).strip():
                facts.listing_probes.append(phrase.group(0).strip())
                break

    # Figure bodies are drawing instructions, not prose; counting them would
    # make the output look truncated whenever a picture is dense.
    prose = FIGURE_RE.sub(" ", body)
    facts.body_chars = len(re.sub(r"\\[a-zA-Z]+|[{}$&\\]", "", prose))
    return facts


# --------------------------------------------------------------------------
# Figure rendering
# --------------------------------------------------------------------------

# tectonic is preferred. It is XeTeX-based, which is what actually builds the
# target paper, so a measured \columnwidth matches the real document instead of
# approximating it; and it fetches the packages a figure needs on demand rather
# than requiring a full TeX distribution. pdflatex stays supported for a machine
# that already has one.
TEX_ENGINES = ("tectonic", "pdflatex")

TEX_ENGINE_HINTS = {
    "tectonic": "brew install tectonic",
    "pdflatex": "install MacTeX or BasicTeX (brew install --cask basictex)",
}


def default_tex_engine() -> str:
    """The first available engine, in preference order."""
    for engine in TEX_ENGINES:
        if shutil.which(engine):
            return engine
    return TEX_ENGINES[0]


def run_tex(engine: str, stem: str, workdir: Path) -> subprocess.CompletedProcess[str]:
    """Compile <stem>.tex inside workdir with the chosen engine.

    tectonic swallows the engine's own chatter unless asked for it, so `--print`
    is what lets a \\typeout reach stdout, and `--keep-logs` leaves behind the
    .log that a failure needs to quote.
    """
    if engine == "tectonic":
        cmd = [resolve_tool("tectonic"), "--print", "--keep-logs", f"{stem}.tex"]
    else:
        cmd = [
            resolve_tool("pdflatex"),
            "-interaction=nonstopmode",
            "-halt-on-error",
            f"{stem}.tex",
        ]
    # argv is built above from a resolved engine and a temp path.
    return subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)  # noqa: S603


def tex_output(result: subprocess.CompletedProcess[str], stem: str, workdir: Path) -> str:
    """Everything the engine said: stdout, stderr, and the .log if one survived."""
    parts = [result.stdout or "", result.stderr or ""]
    log = workdir / f"{stem}.log"
    if log.exists():
        parts.append(log.read_text(encoding="utf-8", errors="replace"))
    return "\n".join(parts)


def harvest_preamble(src: str) -> str:
    """Pull macro and package definitions the figures may depend on."""
    clean = strip_comments(src)
    preamble = clean.split(r"\begin{document}", 1)[0]
    kept = []
    for line in preamble.split("\n"):
        if not PREAMBLE_KEEP_RE.match(line):
            continue
        pkg = re.search(r"\\usepackage(?:\[[^\]]*\])?\{([^}]+)\}", line)
        if pkg and any(p.strip() in PREAMBLE_DROP_PACKAGES for p in pkg.group(1).split(",")):
            continue
        kept.append(line.rstrip())
    return "\n".join(kept)


COLWIDTH_RE = re.compile(r"TEX2GDOC-COLWIDTH=([\d.]+pt)")
SIZE_RE = re.compile(r"TEX2GDOC-SIZE-([a-z]+)=([\d.]+)/([\d.]+pt)")

# The size ladder a class defines. A tikz picture reads these by name, so all of
# them have to come across, not just the current size.
SIZE_STEPS = ("normalsize", "small", "footnotesize", "scriptsize", "tiny")

PROBE_BODY = (
    "\\makeatletter\n"
    "\\typeout{TEX2GDOC-COLWIDTH=\\the\\columnwidth}\n"
    "\\typeout{TEX2GDOC-DIAG textwidth=\\the\\textwidth columnsep=\\the\\columnsep}\n"
    + "".join(
        f"\\{step}\\typeout{{TEX2GDOC-SIZE-{step}=\\f@size/\\f@baselineskip}}\n"
        for step in SIZE_STEPS
    )
    + "\\makeatother\n"
)


@dataclass
class PageMetrics:
    """What the paper's own class says a figure will be typeset into.

    A figure is recompiled outside the paper, so anything the picture reads from
    its surroundings has to be carried across explicitly. Two things are read:
    \\columnwidth, because every axis here is sized `width=\\columnwidth`, and
    the font, because pgfplots sizes its legend box and node text from the font
    while sizing the axis from the width.

    The font has to be carried across as the whole size ladder, not as a single
    `\\fontsize`. A class fixes `\\small`, `\\scriptsize` and the rest at load
    time from its base-size option, and `\\fontsize{9}{11}\\selectfont` does not
    redefine them. This paper styles its legends `font=\\scriptsize`, which is
    6pt in acmart and 7pt in standalone, so selecting 9pt afterwards left the
    legend 17 percent too wide and overlapping the plot.
    """

    column_width: str | None = None
    size_steps: dict[str, tuple[str, str]] = field(default_factory=dict)
    how: str = "not measured"

    @property
    def normal_size(self) -> str | None:
        step = self.size_steps.get("normalsize")
        return step[0] if step else None

    def preamble_lines(self) -> str:
        """The lines to inject into a standalone figure document."""
        out = ""
        for name, (size, skip) in self.size_steps.items():
            out += f"\\renewcommand{{\\{name}}}{{\\fontsize{{{size}}}{{{skip}}}\\selectfont}}\n"
        if self.size_steps:
            out += "\\normalsize\n"
        if self.column_width:
            out += f"\\setlength{{\\columnwidth}}{{{self.column_width}}}\n"
        return out


def measure_page_metrics(src: str, workdir: Path, engine: str) -> PageMetrics:
    """Compile a probe against the paper's own class and read its page metrics.

    Two details, both learned the hard way:

    * A two-column class lays its title block across the full width, so
      \\columnwidth still equals \\textwidth until \\maketitle switches into the
      two-column body. Measuring before that returns a real, plausible, wrong
      number: 506.295pt rather than 241.14749pt for acmart[sigconf], which draws
      every figure at double width. `\\if@twocolumn` is no help, because acmart
      does not set LaTeX's own two-column flag. So run \\maketitle first, and
      fall back to a bare probe only for a class that will not accept one.
    * The values come back through \\typeout rather than \\openout, because
      tectonic does not reliably emit arbitrary auxiliary files while every
      engine prints a \\typeout and writes it to the log.

    `how` is reported, so a fallback never passes silently as a clean measurement.
    """
    m = re.search(r"\\documentclass(\[[^\]]*\])?\{([^}]+)\}", src)
    if not m:
        return PageMetrics(how="no \\documentclass found")
    stem = "probe"
    declaration = f"\\documentclass{m.group(1) or ''}{{{m.group(2)}}}\n"
    attempts = [
        (
            "after \\maketitle",
            declaration + "\\begin{document}\n"
            "\\title{probe}\\author{probe}\n\\maketitle\n" + PROBE_BODY + "\\end{document}\n",
        ),
        (
            "bare document, no \\maketitle",
            declaration + "\\begin{document}\n" + PROBE_BODY + "x\\end{document}\n",
        ),
    ]
    for how, source in attempts:
        # Clear the previous attempt's log first: tex_output() reads it, and a
        # stale one could hand back the earlier attempt's numbers as this one's.
        (workdir / f"{stem}.log").unlink(missing_ok=True)
        (workdir / f"{stem}.tex").write_text(source, encoding="utf-8")
        output = tex_output(run_tex(engine, stem, workdir), stem, workdir)
        width = COLWIDTH_RE.search(output)
        if width:
            return PageMetrics(
                column_width=width.group(1),
                size_steps={m.group(1): (m.group(2), m.group(3)) for m in SIZE_RE.finditer(output)},
                how=how,
            )
    return PageMetrics(how="probe did not compile")


def render_figure(
    fig: Figure, preamble: str, metrics: PageMetrics, workdir: Path, dpi: int, engine: str
) -> tuple[bool, str]:
    """Compile one TikZ picture standalone and rasterize it. Returns (ok, detail)."""
    stem = f"fig-{fig.label}"

    def build(pre: str) -> subprocess.CompletedProcess[str]:
        (workdir / f"{stem}.tex").write_text(
            "\\documentclass[tikz,border=2pt]{standalone}\n"
            "\\usepackage{amsmath,amssymb}\n"
            f"{pre}\n"
            "\\usepackage{tikz}\n"
            "\\begin{document}\n"
            f"{metrics.preamble_lines()}"
            f"{fig.tikz}\n"
            "\\end{document}\n",
            encoding="utf-8",
        )
        return run_tex(engine, stem, workdir)

    result = build(preamble)
    if result.returncode != 0:
        # The paper's preamble may carry class-specific macros that standalone
        # rejects. Retry bare, and say so rather than passing it off as clean.
        result = build("\\usepackage{pgfplots}\n\\pgfplotsset{compat=newest}")
        if result.returncode != 0:
            tail = tex_output(result, stem, workdir)[-1200:]
            return False, f"{engine} failed for {fig.label}:\n{tail}"
        fig.fallback_preamble = True

    subprocess.run(  # noqa: S603 - fixed argv, shell=False, arguments built here
        [resolve_tool("pdftocairo"), "-png", "-r", str(dpi), "-singlefile", f"{stem}.pdf", stem],
        cwd=workdir,
        capture_output=True,
        text=True,
    )
    png = workdir / f"{stem}.png"
    if not png.exists() or png.stat().st_size < 2000:
        return False, f"rasterization produced no usable PNG for {fig.label}"
    fig.png = png
    detail = "rendered with fallback preamble" if fig.fallback_preamble else "rendered"
    return True, f"{detail} ({png.stat().st_size // 1024} KB)"


# --------------------------------------------------------------------------
# Source rewriting
# --------------------------------------------------------------------------


def flatten_nested_tabulars(src: str) -> tuple[str, int]:
    """Collapse tabulars nested inside table cells into single-line text.

    Pandoc's LaTeX reader cannot parse a tabular inside a cell and discards the
    entire enclosing table without warning. Flattening keeps the content and
    loses only the line breaks within that cell.
    """
    count = 0
    while True:
        nested = [s for s in find_tabular_spans(src) if s[2] > 0]
        if not nested:
            return src, count
        # Innermost first: the deepest span with no tabular inside it.
        nested.sort(key=lambda s: (-s[2], s[0]))
        start, end, _ = nested[0]
        block = src[start:end]
        # The column spec contains braces of its own (`{@{}l@{}}`), so it has to
        # be consumed by brace matching; a non-brace character class stops at the
        # first inner `}` and leaves debris such as `l@` in the output.
        head = re.match(r"\\begin\{tabular\}(\[[^\]]*\])?\s*\{", block)
        inner = block
        if head:
            spec_open = head.end() - 1
            spec = extract_braced(block, spec_open)
            inner = block[spec_open + len(spec) + 2 :]
        inner = re.sub(r"\\end\{tabular\}$", "", inner)
        parts = [p.strip() for p in inner.split("\\\\") if p.strip()]
        src = src[:start] + "; ".join(parts) + src[end:]
        count += 1


ALLTT_RE = re.compile(r"\\begin\{alltt\}(.*?)\\end\{alltt\}", re.DOTALL)
MATH_ESCAPE_RE = re.compile(r"\\\((.*?)\\\)", re.DOTALL)
STARRED_ENV_RE = re.compile(r"\\(begin|end)\{(table|figure)\*\}")
REF_RE = re.compile(r"\\ref\{((?:tab|fig|thm):[^}]+)\}")

# Symbols that appear inside verbatim-style Lean listings. A listing is code,
# so these belong in it as characters, not as equation objects.
MATH_TO_UNICODE = {
    r"\mathbb{R}": "ℝ",
    r"\mathbb{N}": "ℕ",
    r"\mathbb{Z}": "ℤ",
    r"\mathbb{Q}": "ℚ",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\in": "∈",
    r"\notin": "∉",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\ne": "≠",
    r"\neq": "≠",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\mapsto": "↦",
    r"\implies": "⟹",
    r"\wedge": "∧",
    r"\land": "∧",
    r"\vee": "∨",
    r"\lor": "∨",
    r"\lnot": "¬",
    r"\neg": "¬",
    r"\Lambda": "Λ",
    r"\lambda": "λ",
    r"\Theta": "Θ",
    r"\theta": "θ",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ε",
    r"\varepsilon": "ε",
    r"\mu": "μ",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\Sigma": "Σ",
    r"\Pi": "Π",
    r"\times": "×",
    r"\cdot": "·",
    r"\circ": "∘",
    r"\subseteq": "⊆",
    r"\cup": "∪",
    r"\cap": "∩",
    r"\emptyset": "∅",
    r"\infty": "∞",
}


def alltt_to_verbatim(src: str) -> tuple[str, int, set[str]]:
    """Turn `alltt` listings into `verbatim`, with math escapes as characters.

    Pandoc reads `alltt` as a paragraph, so every run of spaces collapses and
    the indentation of a code listing is lost: `structure Loop where` arrives
    as `structureLoopwhere`. A `verbatim` block keeps spacing.
    The `\\(...\\)` escapes inside it have to become literal characters first,
    since verbatim would otherwise print the macro names.

    Returns the rewritten source, the number of blocks converted, and any
    macros with no character mapping — those are reported, never dropped
    silently.
    """
    unmapped: set[str] = set()

    def convert_math(m: re.Match[str]) -> str:
        body = m.group(1).strip()
        if body in MATH_TO_UNICODE:
            return MATH_TO_UNICODE[body]
        for macro, char in MATH_TO_UNICODE.items():
            body = body.replace(macro, char)
        leftover = re.findall(r"\\[a-zA-Z]+", body)
        if leftover:
            unmapped.update(leftover)
        return re.sub(r"[{}]", "", body)

    count = 0

    def convert_block(m: re.Match[str]) -> str:
        nonlocal count
        count += 1
        body = MATH_ESCAPE_RE.sub(convert_math, m.group(1))
        body = re.sub(r"^\\[a-zA-Z]+\s*", "", body)  # a size command after \begin{alltt}
        body = body.replace(r"\{", "{").replace(r"\}", "}").replace(r"\\", "\\")
        return "\\begin{verbatim}" + body.strip("\n") + "\n\\end{verbatim}"

    return ALLTT_RE.sub(convert_block, src), count, unmapped


def unstar_float_environments(src: str) -> tuple[str, int]:
    """Convert `table*`/`figure*` to their unstarred forms.

    Pandoc keeps the caption of a `table` and drops the caption of a `table*`
    without warning. Column spanning has no meaning in a single-column .docx,
    so unstarring costs nothing and returns the captions.
    """
    src, n = STARRED_ENV_RE.subn(lambda m: f"\\{m.group(1)}{{{m.group(2)}}}", src)
    return src, n // 2


def resolve_float_references(src: str) -> tuple[str, int]:
    """Replace `\\ref` to a table, figure, or theorem with its number.

    Pandoc numbers sections but not floats, so these otherwise reach the reader
    as raw keys such as `[tab:map]`. Numbering follows the order the labels
    appear in the source, which is the order the reader sees.
    """
    numbering: dict[str, int] = {}
    counters: dict[str, int] = {}
    for m in re.finditer(r"\\label\{((?:tab|fig|thm):[^}]+)\}", src):
        key = m.group(1)
        if key in numbering:
            continue
        kind = key.split(":")[0]
        counters[kind] = counters.get(kind, 0) + 1
        numbering[key] = counters[kind]
    if not numbering:
        return src, 0
    replaced = 0

    def sub(m: re.Match[str]) -> str:
        nonlocal replaced
        key = m.group(1)
        if key not in numbering:
            return m.group(0)
        replaced += 1
        return str(numbering[key])

    return REF_RE.sub(sub, src), replaced


@dataclass(frozen=True)
class Rewrites:
    """What each source transformation did, for the run to report.

    Four counts and a set of macro names. As a dict[str, object] every read
    needed a cast, and `sorted(stats.unmapped)` type-checked only because
    `object` says nothing.
    """

    flattened: int
    unstarred: int
    refs: int
    listings: int
    unmapped: set[str]


def rewrite_source(src: str, figures: list[Figure], figdir: Path) -> tuple[str, Rewrites]:
    """Apply every transformation pandoc needs, and report what each one did."""
    by_block = {f.block: f for f in figures if f.png}

    def replace(match: re.Match[str]) -> str:
        fig = by_block.get(match.group(0))
        # by_block only holds figures with a rendered png, so the second test is
        # unreachable; it is a check rather than an assert because asserts are
        # stripped under -O, and it narrows the type for free.
        if fig is None or fig.png is None:
            return match.group(0)
        include = f"\\includegraphics[width=\\linewidth]{{{figdir.name}/{fig.png.name}}}"
        return TIKZ_RE.sub(lambda _: include, match.group(0), count=1)

    src = FIGURE_RE.sub(replace, src)
    src, flattened = flatten_nested_tabulars(src)
    src, unstarred = unstar_float_environments(src)
    src, refs = resolve_float_references(src)
    src, listings, unmapped = alltt_to_verbatim(src)
    return src, Rewrites(
        flattened=flattened,
        unstarred=unstarred,
        refs=refs,
        listings=listings,
        unmapped=unmapped,
    )


# --------------------------------------------------------------------------
# Presentation: page geometry, styles, table widths, code colouring
# --------------------------------------------------------------------------

# Twentieths of a point (DXA). US Letter with one-inch margins, which is what
# Google Docs creates by default. Narrower margins would give the wide tables
# more room, but a reviewer opening this next to a document they made in Docs
# should not find the page set up differently.
PAGE_WIDTH = 12240
PAGE_HEIGHT = 15840
MARGIN = 1440  # 1 inch
TEXT_WIDTH = PAGE_WIDTH - 2 * MARGIN

# Google Docs renders percentage table widths unreliably, so every width this
# script writes is absolute DXA. Courier New is the one monospace face present
# on Windows, macOS, and Google Docs alike; a prettier choice would silently
# fall back to a proportional font for some readers.
CODE_FONT = "Courier New"
# The face a renderer lays equations out in. The document never names it;
# this is only the hint that survives in fontTable.
MATH_FONT = "Cambria Math"
FONT_NAME_SLOTS = (Attr.ASCII, Attr.HIGH_ANSI, Attr.COMPLEX_SCRIPT, Attr.EAST_ASIAN)


def section_properties() -> SafeElement:
    """The page size and margins, built rather than spelled out as markup.

    A function rather than a module constant because an element is mutable and
    is consumed by being appended to a tree: handing the same one to two
    documents would move it out of the first.
    """
    return make(
        Tag.SECTION_PROPERTIES,
        children=[
            make(Tag.PAGE_SIZE, Attr.WIDTH.of(PAGE_WIDTH), Attr.HEIGHT.of(PAGE_HEIGHT)),
            make(
                Tag.PAGE_MARGIN,
                Attr.TOP.of(MARGIN),
                Attr.RIGHT.of(MARGIN),
                Attr.BOTTOM.of(MARGIN),
                Attr.LEFT.of(MARGIN),
                Attr.HEADER.of(HEADER_FOOTER_MARGIN),
                Attr.FOOTER.of(HEADER_FOOTER_MARGIN),
                Attr.GUTTER.of(0),
            ),
        ],
    )


# Only what is functional. Everything else keeps pandoc's own definition, so
# Word and Google Docs map Heading1-3, Title, Caption and the rest onto native
# styles, and their style menus keep working on them.
#
# An earlier version restyled all of these with custom colours, sizes and
# spacing. It looked deliberate and it made the imported document hard to work
# with: Docs honours a built-in style's overridden definition, so applying
# "Heading 2" from the menu reproduced this script's blue rather than the
# document's own. Overriding the appearance of a built-in style is the thing to
# avoid; setting layout that the reader cannot otherwise get is not.
@dataclass(frozen=True)
class StylePatch:
    """What a style must carry, as properties rather than as markup.

    Every field is a layout decision someone can argue with. Written as XML
    these were a wall of angle brackets in which `w:line="276"` and
    `w:lineRule="auto"` had to travel together and nothing said so; here the
    pairing is the builder's problem and a reader sees only the decision.
    """

    space_after: int | None = None
    space_before: int | None = None
    line: int | None = None
    # "auto" makes the line as tall as the tallest thing on it, so any line
    # carrying an inline equation grows and the page develops ragged leading
    # that LaTeX never shows. "exact" pins it, which is what makes the two
    # match. See TALL_MATH for what has to be exempted.
    line_rule: Value = Value.EXACT
    keep_next: bool = False
    keep_lines: bool = False
    centered: bool = False
    shading: str | None = None
    monospace: bool = False

    def paragraph_properties(self) -> SafeElement | None:
        properties: list[SafeElement] = []
        if self.keep_next:
            properties.append(make(Tag.KEEP_NEXT))
        if self.keep_lines:
            properties.append(make(Tag.KEEP_LINES))
        spacing: list[Attribute] = []
        if self.space_before is not None:
            spacing.append(Attr.BEFORE.of(self.space_before))
        if self.space_after is not None:
            spacing.append(Attr.AFTER.of(self.space_after))
        if self.line is not None:
            # A line value means nothing without its rule, so the builder
            # supplies the pair and no caller can forget half of it.
            spacing += [Attr.LINE.of(self.line), Attr.LINE_RULE.of(self.line_rule)]
        if spacing:
            properties.append(make(Tag.SPACING, *spacing))
        if self.shading is not None:
            properties.append(
                make(
                    Tag.SHADING,
                    Attr.VAL.of(Value.CLEAR),
                    Attr.COLOR.of(Value.AUTO),
                    Attr.FILL.of(self.shading),
                )
            )
        if self.centered:
            properties.append(make(Tag.JUSTIFICATION, Attr.VAL.of(Value.CENTER)))
        return make(Tag.PARAGRAPH_PROPERTIES, children=properties) if properties else None

    def run_properties(self) -> SafeElement | None:
        if not self.monospace:
            return None
        fonts = make(
            Tag.FONTS,
            Attr.ASCII.of(CODE_FONT),
            Attr.HIGH_ANSI.of(CODE_FONT),
            Attr.COMPLEX_SCRIPT.of(CODE_FONT),
        )
        return make(Tag.RUN_PROPERTIES, children=[fonts])

    def elements(self) -> list[SafeElement]:
        return [e for e in (self.paragraph_properties(), self.run_properties()) if e is not None]


# 1.15 line spacing, expressed the way OOXML counts it: 240 twentieths of a
# point is single, so 276 is 1.15.
LINE_115 = 276

STYLE_PATCHES: dict[StyleId, StylePatch] = {
    # 1.15 line and 8pt after a paragraph are Google Docs' own body defaults, so
    # the import matches a document created there rather than announcing itself.
    # space_before is pinned to 0 rather than left to inherit, because that is
    # Docs' own default and a host template with a nonzero before would
    # otherwise reflow the body. Dropping it was an unintended behaviour change
    # when these patches moved from literal XML to StylePatch on 2026-08-10.
    StyleId.NORMAL: StylePatch(space_before=0, space_after=160, line=LINE_115),
    StyleId.BODY_TEXT: StylePatch(space_before=0, space_after=160, line=LINE_115),
    # List items stay tight: 8pt between consecutive bullets reads as broken.
    StyleId.COMPACT: StylePatch(space_before=0, space_after=0, line=LINE_115),
    # Centring a figure is layout, not decoration, and `verify()` checks for it.
    StyleId.CAPTIONED_FIGURE: StylePatch(keep_next=True, centered=True),
    StyleId.FIGURE: StylePatch(keep_next=True, centered=True),
    StyleId.IMAGE_CAPTION: StylePatch(centered=True),
    # Code has to be monospace to be readable at all, and the shaded block is
    # what separates a listing from prose.
    StyleId.VERBATIM_CHAR: StylePatch(monospace=True),
    StyleId.SOURCE_CODE: StylePatch(
        keep_lines=True, space_before=80, space_after=80, line=240, shading="F7F8FA", monospace=True
    ),
}


PATCH_BY_STYLE_ID: Final[dict[str, StylePatch]] = {
    style.value: patch for style, patch in STYLE_PATCHES.items()
}


def patch_styles(styles_xml: str) -> str:
    """Replace the paragraph and run properties of the styles listed above.

    Nothing else in styles.xml is touched. The document default font size, every
    heading definition and pandoc's own table style are left exactly as pandoc
    wrote them, because those are what Google Docs reads to decide what its own
    named styles mean.
    """
    root = parse_ooxml(styles_xml, "word/styles.xml")
    for style in descendants(root, Tag.STYLE):
        style_id = attribute(style, Attr.STYLE_ID)
        patch = PATCH_BY_STYLE_ID.get(style_id) if style_id else None
        if patch is None:
            continue
        for tag in (Tag.PARAGRAPH_PROPERTIES, Tag.RUN_PROPERTIES):
            for existing in children(style, tag):
                style.remove(existing)
        style.extend(patch.elements())
    return serialize_ooxml(root)


def build_reference_doc(workdir: Path) -> Path | None:
    """Write a pandoc reference document carrying the page geometry and styles.

    Derived from pandoc's own default rather than kept as a stored file, so it
    always matches the installed pandoc instead of drifting from it.
    """
    default = subprocess.run(  # noqa: S603 - fixed argv, shell=False
        [resolve_tool("pandoc"), "--print-default-data-file", "reference.docx"],
        capture_output=True,
    )
    if default.returncode != 0 or not default.stdout:
        return None
    src = workdir / "reference-default.docx"
    src.write_bytes(default.stdout)
    out = workdir / "reference.docx"

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                data = patch_styles(data.decode("utf-8")).encode("utf-8")
            elif item.filename == "word/document.xml":
                root = parse_ooxml(data.decode("utf-8"), item.filename)
                body = child(root, Tag.BODY)
                if body is not None:
                    for existing in children(body, Tag.SECTION_PROPERTIES):
                        body.remove(existing)
                    body.append(section_properties())
                data = serialize_ooxml(root).encode("utf-8")
            zout.writestr(item, data)
    return out


# Lean is not one of the languages pandoc can highlight, so the colouring is
# applied here, as direct run formatting. Direct formatting is also what
# survives an import most reliably.
LEAN_KEYWORDS = {
    "theorem",
    "lemma",
    "def",
    "abbrev",
    "structure",
    "class",
    "instance",
    "inductive",
    "extends",
    "where",
    "with",
    "by",
    "fun",
    "let",
    "have",
    "show",
    "from",
    "match",
    "do",
    "if",
    "then",
    "else",
    "open",
    "namespace",
    "end",
    "variable",
    "example",
    "noncomputable",
    "private",
    "protected",
    "partial",
    "mutual",
    "deriving",
    "at",
}
LEAN_TYPES = {
    "Prop",
    "Type",
    "Sort",
    "ℝ",
    "ℕ",
    "ℤ",
    "ℚ",
    "Bool",
    "Nat",
    "Real",
    "Set",
    "Finset",
    "PMF",
}


class TokenKind(StrEnum):
    """What a token in a code listing is, and the colour that says so.

    The colour travels with the kind rather than sitting in a parallel dict:
    adding a kind and forgetting its colour used to be a `KeyError` at render
    time, and is now impossible to write.
    """

    KEYWORD = "keyword"
    TYPE = "type"
    NUMBER = "number"
    SYMBOL = "symbol"
    COMMENT = "comment"
    PLAIN = "plain"

    @property
    def color(self) -> str:
        match self:
            case TokenKind.KEYWORD:
                return "0033B3"
            case TokenKind.TYPE:
                return "0F7B6C"
            case TokenKind.NUMBER:
                return "B26B00"
            case TokenKind.SYMBOL:
                return "7A3E9D"
            case TokenKind.COMMENT:
                return "6A737D"
            case TokenKind.PLAIN:
                return "1A1A1A"


TOKEN_RE = re.compile(r"--[^\n]*|[A-Za-zℝℕℤℚ_][A-Za-z0-9ℝℕℤℚ_.']*|\d+|\s+|.")


def classify_token(tok: str) -> TokenKind:
    if tok.startswith("--"):
        return TokenKind.COMMENT
    if tok in LEAN_KEYWORDS:
        return TokenKind.KEYWORD
    if tok in LEAN_TYPES or (tok[:1].isupper() and tok[:1].isalpha()):
        return TokenKind.TYPE
    if tok.isdigit():
        return TokenKind.NUMBER
    if not tok.strip():
        return TokenKind.PLAIN
    if not (tok[0].isalnum() or tok[0] == "_"):
        return TokenKind.SYMBOL
    return TokenKind.PLAIN


def make_code_run(token: str) -> SafeElement:
    """One coloured run for one token."""
    kind = classify_token(token)
    properties = [
        make(Tag.RUN_STYLE, Attr.VAL.of(StyleId.VERBATIM_CHAR)),
        make(Tag.COLOR, Attr.VAL.of(kind.color)),
    ]
    if kind is TokenKind.KEYWORD:
        properties.append(make(Tag.BOLD))
    elif kind is TokenKind.COMMENT:
        properties.append(make(Tag.ITALIC))
    return make(
        Tag.RUN,
        children=[
            make(Tag.RUN_PROPERTIES, children=properties),
            make(Tag.TEXT, XML_SPACE_PRESERVE, text=token),
        ],
    )


def highlight_code_paragraphs(root: SafeElement) -> int:
    """Recolour the runs of every `SourceCode` paragraph, token by token.

    A run carrying no `w:t` is passed through untouched. Pandoc puts the line
    break between listing lines in a run of its own, so dropping those would run
    the whole listing onto one line.
    """
    count = 0
    for para in descendants(root, Tag.PARAGRAPH):
        ppr = child(para, Tag.PARAGRAPH_PROPERTIES)
        if ppr is None or value_of(child(ppr, Tag.PARAGRAPH_STYLE)) != StyleId.SOURCE_CODE:
            continue
        count += 1
        rebuilt: list[SafeElement] = []
        for node in (SafeElement(e) for e in para):
            text = child(node, Tag.TEXT) if node.tag == qn(Tag.RUN) else None
            if text is None:
                rebuilt.append(node)
                continue
            rebuilt.extend(make_code_run(tok) for tok in TOKEN_RE.findall(text.text or ""))
        para[:] = rebuilt
    return count


def cell_text(cell: SafeElement) -> str:
    return " ".join("".join(t for t in cell.itertext() if isinstance(t, str)).split())


def cell_is_monospace(cell: SafeElement) -> bool:
    return any(value_of(s) == "VerbatimChar" for s in descendants(cell, Tag.RUN_STYLE))


def column_widths(rows: list[list[SafeElement]], columns: int) -> list[int]:
    """Size columns by what they hold, then scale the row to the full text width.

    A column's floor is set by its longest unbreakable word, not by a flat
    percentage: a header such as `Provenance` or an identifier such as
    `unique_fixed_point_of_monotone` wraps mid-word once the column is
    narrower than the word itself. Whatever space is left over is then shared
    out by how much text each column actually carries, so prose columns get the
    room and a column of row numbers does not.
    """
    # Approximate advance width per character at the body size, in DXA.
    CHAR_PROPORTIONAL = 108
    CHAR_MONOSPACE = 124
    CELL_PADDING = 200

    floors, weights = [], []
    for i in range(columns):
        lengths, longest_token, mono = [], 1, 0
        for row in rows:
            if i >= len(row):
                continue
            text = cell_text(row[i])
            lengths.append(len(text))
            for word in text.split():
                longest_token = max(longest_token, len(word))
            if cell_is_monospace(row[i]):
                mono += 1
        if not lengths:
            floors.append(CELL_PADDING)
            weights.append(1.0)
            continue
        per_char = CHAR_MONOSPACE if mono > len(rows) * 0.5 else CHAR_PROPORTIONAL
        floor = min(longest_token * per_char + CELL_PADDING, int(TEXT_WIDTH * 0.34))
        lengths.sort()
        typical = lengths[int(len(lengths) * 0.75)] if lengths else 1
        floors.append(floor)
        weights.append(max(float(typical), 1.0))

    total_floor = sum(floors)
    if total_floor >= TEXT_WIDTH:
        # Not enough room for every longest word; shrink in proportion so the
        # loss is shared rather than dumped on one column.
        scale = TEXT_WIDTH / total_floor
        widths = [int(f * scale) for f in floors]
    else:
        surplus = TEXT_WIDTH - total_floor
        weight_total = sum(weights)
        # strict=True: floors and weights are built in the same loop, so a
        # length mismatch would be a bug rather than something to absorb.
        widths = [int(f + surplus * n / weight_total) for f, n in zip(floors, weights, strict=True)]
    widths[-1] += TEXT_WIDTH - sum(widths)  # absorb rounding, so the row is exact
    return widths


def _replace_child(
    parent: SafeElement, tag: Tag, *attributes: Attribute, first: bool = False
) -> None:
    """Drop every existing child with this tag and add one with these attributes."""
    for existing in children(parent, tag):
        parent.remove(existing)
    element = make(tag, *attributes)
    if first:
        parent.insert(0, element)
    else:
        parent.append(element)


def resize_tables(root: SafeElement) -> int:
    """Make every table span the text width, with columns sized to their content."""
    count = 0
    for tbl in descendants(root, Tag.TABLE):
        rows = [children(tr, Tag.TABLE_CELL) for tr in children(tbl, Tag.TABLE_ROW)]
        if not rows:
            continue
        widths = column_widths(rows, max(len(r) for r in rows))
        count += 1

        grid = child(tbl, Tag.TABLE_GRID)
        if grid is not None:
            tail = grid.tail
            grid.clear()
            grid.tail = tail
            for width in widths:
                grid.append(make(Tag.GRID_COLUMN, Attr.WIDTH.of(width)))

        tblpr = child(tbl, Tag.TABLE_PROPERTIES)
        if tblpr is None:
            tblpr = make(Tag.TABLE_PROPERTIES)
            tbl.insert(0, tblpr)
        _replace_child(tblpr, Tag.TABLE_WIDTH, Attr.TYPE.of(Value.DXA), Attr.WIDTH.of(TEXT_WIDTH))
        _replace_child(tblpr, Tag.TABLE_LAYOUT, Attr.TYPE.of(Value.FIXED))

        # Every cell carries its own width as well as the grid: Google Docs
        # honours the cell width, and a table with only a grid comes in squeezed.
        for tr in children(tbl, Tag.TABLE_ROW):
            index = 0
            for tc in children(tr, Tag.TABLE_CELL):
                tcpr = child(tc, Tag.CELL_PROPERTIES)
                if tcpr is None:
                    tcpr = make(Tag.CELL_PROPERTIES)
                    tc.insert(0, tcpr)
                span = int(value_of(child(tcpr, Tag.GRID_SPAN)) or 1)
                width = sum(widths[index : index + span]) or widths[-1]
                index += span
                _replace_child(
                    tcpr, Tag.CELL_WIDTH, Attr.TYPE.of(Value.DXA), Attr.WIDTH.of(width), first=True
                )
    return count


# Style ids Google Docs has no native equivalent for, but which this document
# uses for ordinary running prose. Docs flattens an unmapped style into direct
# formatting on every paragraph that used it, and direct formatting is
# unreachable from its style menu, so the body of an imported paper cannot be
# restyled at all. Both resolve to exactly what `Normal` resolves to here
# (`BodyText` carries the same spacing, `FirstParagraph` carries none and
# inherits), so the remap is invisible on the page and total in the menu.
# Where each of pandoc's paragraph styles has to land for Google Docs to map it
# onto something in its own style menu. Docs offers Normal text, Title, Subtitle
# and Heading 1-6 and nothing else; every other style it meets becomes direct
# formatting on the paragraph, which the menu cannot reach. A paper imported
# without this can have its headings restyled and nothing else.
#
# Compact is table-cell text and differs from Normal only by carrying no space
# after, so the remap pins that one property back on the paragraph and lets
# everything else follow the style. Captions and SourceCode keep their own
# styles: centring and the monospace block are layout this document has to
# state, and there is no native style that carries either.
STYLE_REMAP: dict[StyleId, StyleId] = {
    StyleId.BODY_TEXT: StyleId.NORMAL,
    StyleId.FIRST_PARAGRAPH: StyleId.NORMAL,
    StyleId.COMPACT: StyleId.NORMAL,
    StyleId.BIBLIOGRAPHY: StyleId.NORMAL,
}
# Abstract, BlockText and Author are deliberately absent. Abstract carries
# keepNext, keepLines and its own spacing; BlockText carries the left and right
# indents that make a block quote a block quote; and Author would have to go to
# Subtitle, which in pandoc's styles carries a numPr. Remapping any of them
# would move something on the page, which is the one thing this pass promises
# not to do. Together they are three paragraphs in a 667-paragraph paper.

# OMML constructs taller than a line of text. A paragraph holding one of these
# keeps "auto" leading, because a fixed line height crops what overflows it and
# a clipped integral sign is worse than an uneven page.
TALL_MATH = (
    Tag.NARY,
    Tag.FRACTION,
    Tag.RADICAL,
    Tag.MATRIX,
    Tag.SUB_SUPERSCRIPT,
    Tag.LIMIT_LOWER,
    Tag.LIMIT_UPPER,
    Tag.BOX,
)


# The CT_PPr members that come before w:spacing in the schema sequence. Only
# the ones pandoc or this converter can emit are listed; anything unlisted
# sorts after spacing, which is the safe direction to guess in.
PRECEDE_SPACING = frozenset(
    qn(tag)
    for tag in (
        Tag.PARAGRAPH_STYLE,
        Tag.KEEP_NEXT,
        Tag.KEEP_LINES,
        Tag.NUMBERING_PROPERTIES,
        Tag.SHADING,
    )
)

REMAP_BY_STYLE_ID: Final[dict[str, StyleId]] = {
    style.value: target for style, target in STYLE_REMAP.items()
}


def paragraph_properties(paragraph: SafeElement) -> SafeElement:
    """The paragraph's `w:pPr`, created at the front if it has none."""
    existing = child(paragraph, Tag.PARAGRAPH_PROPERTIES)
    if existing is not None:
        return existing
    created = make(Tag.PARAGRAPH_PROPERTIES)
    paragraph.insert(0, created)
    return created


def set_spacing(paragraph: SafeElement, *attributes: Attribute) -> None:
    """Set attributes on the paragraph's own `w:spacing`, creating it if needed.

    Placed before `w:ind`, `w:jc`, `w:rPr` and `w:sectPr` when any of them are
    present, because CT_PPr is a sequence and those four follow spacing in it.
    """
    ppr = paragraph_properties(paragraph)
    spacing = child(ppr, Tag.SPACING)
    if spacing is None:
        spacing = make(Tag.SPACING)
        # Positioned from what precedes spacing rather than from what follows
        # it: the preceding set is short and closed, while a dozen elements may
        # follow, and naming only some of those put spacing out of order
        # whenever a paragraph carried one of the rest.
        last = -1
        for index, element in enumerate(ppr):
            if element.tag in PRECEDE_SPACING:
                last = index
        ppr.insert(last + 1, spacing)
    for attribute_ in attributes:
        spacing.set(qn(attribute_.name), attribute_.value)


def map_styles_for_docs(document: SafeElement) -> int:
    """Point every paragraph at a style Google Docs knows, where one exists."""
    remapped = 0
    for pstyle in descendants(document, Tag.PARAGRAPH_STYLE):
        current = value_of(pstyle)
        # A StrEnum member hashes as its value, so the plain string keys the
        # dict directly: no membership test and no StyleId() construction.
        target = REMAP_BY_STYLE_ID.get(current) if current else None
        if target is None:
            continue
        pstyle.set(qn(Attr.VAL), target)
        remapped += 1
        if current == StyleId.COMPACT:
            # pStyle sits inside pPr, which sits inside the paragraph.
            ppr = pstyle.getparent()
            paragraph = None if ppr is None else ppr.getparent()
            if paragraph is not None:
                set_spacing(SafeElement(paragraph), Attr.AFTER.of(0))
    return remapped


def exempt_tall_math_from_fixed_leading(document: SafeElement) -> int:
    """Give paragraphs holding a tall equation their leading back.

    Everything else runs on the fixed line height that makes the page match
    LaTeX. These are the paragraphs where that would crop the maths instead.
    """
    exempted = 0
    for paragraph in descendants(document, Tag.PARAGRAPH):
        tags = {e.tag for e in paragraph.iter()}
        if any(qn(tall) in tags for tall in TALL_MATH):
            set_spacing(SafeElement(paragraph), Attr.LINE_RULE.of(Value.AUTO))
            exempted += 1
    return exempted


def drop_empty_comments_part(parts: dict[str, bytes]) -> bool:
    """Remove word/comments.xml when it holds no comments.

    pandoc's reference document ships the part whether or not anything uses it,
    and this converter never produces a comment. An empty part is inert, but it
    is also a part every consumer has to open and a relationship every consumer
    has to resolve, for nothing.
    """
    part = "word/comments.xml"
    if part not in parts:
        return False
    if descendants(parse_ooxml(parts[part].decode("utf-8"), part), Tag.COMMENT):
        return False
    del parts[part]

    types = "[Content_Types].xml"
    if types in parts:
        root = parse_ooxml(parts[types].decode("utf-8"), types)
        for override in list(root):
            if attribute(SafeElement(override), Attr.PART_NAME) == f"/{part}":
                root.remove(override)
        parts[types] = serialize_ooxml(root).encode("utf-8")

    rels = "word/_rels/document.xml.rels"
    if rels in parts:
        root = parse_ooxml(parts[rels].decode("utf-8"), rels)
        for rel in descendants(root, Tag.RELATIONSHIP):
            if (attribute(rel, Attr.TARGET) or "").endswith("comments.xml"):
                root.remove(rel)
        parts[rels] = serialize_ooxml(root).encode("utf-8")
    return True


def clear_theme_fonts(parts: dict[str, bytes]) -> list[str]:
    """Leave the theme naming no typeface at all, for any script.

    pandoc's reference doc ships the Microsoft 365 theme, whose faces are Aptos
    and Aptos Display. Nothing outside a current Office install has them, so
    every other reader gets a substitution nobody chose.

    Naming a safer font would only move the problem: any name is still this
    document overriding the template it was imported into. An empty `typeface`
    is the OOXML way to say nothing, and it leaves headings and body text
    resolving to whatever Word, Docs or LibreOffice calls normal text.

    The whole font scheme is swept, not just the latin slot. A theme carries a
    latin, an east-Asian and a complex-script face per slot, and then a
    per-script table: 94 more names in this one, for Japanese, Korean, Chinese,
    Arabic, Hebrew and the rest. Clearing two of them and calling the file
    font-free is the kind of claim that survives only because nobody looked at
    the other 94. PANOSE goes too: it is a ten-byte description of the face's
    own metrics, and a renderer that cannot match a name falls back to matching
    those, which lands on something Aptos-shaped again.

    The styles need no edit. They already point at the theme
    (`w:asciiTheme="minorHAnsi"`) rather than at a literal name, and the body
    pins no font at all, so clearing the scheme re-points the whole document.
    """
    theme = "word/theme/theme1.xml"
    if theme not in parts:
        return []
    root = parse_ooxml(parts[theme].decode("utf-8"), theme)
    cleared: list[str] = []
    for slot in (Tag.MAJOR_FONT, Tag.MINOR_FONT):
        for scheme in descendants(root, slot):
            # Every descendant, so the per-script table is covered along with
            # the three named slots.
            for element in (SafeElement(e) for e in scheme.iter()):
                named = attribute(element, Attr.TYPEFACE)
                if named:
                    cleared.append(named)
                if named is not None:
                    element.set(qn(Attr.TYPEFACE), "")
                if attribute(element, Attr.PANOSE) is not None:
                    element.set(qn(Attr.PANOSE), "")
    parts[theme] = serialize_ooxml(root).encode("utf-8")
    return cleared


def strip_unused_font_table(parts: dict[str, bytes]) -> list[str]:
    """Drop fontTable entries for faces the document no longer mentions.

    The table is advisory: it carries metrics and embedding hints for fonts the
    document uses, and Word rebuilds it on save. After the theme is cleared,
    every entry pandoc inherited from its reference doc describes a face
    nothing asks for, and Aptos sitting in the table is still this file naming
    Aptos to whoever opens it.

    Two survive on purpose. Whatever the document pins literally, which is the
    monospace face for code listings, and Cambria Math, which is the face a
    renderer reaches for when it lays out the equations: the document names no
    maths font itself, so removing its table entry would be removing the one
    hint about what the maths is meant to look like.
    """
    table = "word/fontTable.xml"
    if table not in parts:
        return []
    referenced = {MATH_FONT}
    for part in ("word/document.xml", "word/styles.xml"):
        if part not in parts:
            continue
        tree = parse_ooxml(parts[part].decode("utf-8"), part)
        for fonts in descendants(tree, Tag.FONTS):
            referenced |= {name for slot in FONT_NAME_SLOTS if (name := attribute(fonts, slot))}

    root = parse_ooxml(parts[table].decode("utf-8"), table)
    dropped: list[str] = []
    for font in children(root, Tag.FONT):
        name = attribute(font, Attr.NAME)
        if name is not None and name not in referenced:
            dropped.append(name)
            root.remove(font)
    parts[table] = serialize_ooxml(root).encode("utf-8")
    return dropped


def strip_font_embedding(parts: dict[str, bytes]) -> int:
    """Remove font-embedding directives inherited from pandoc's reference doc.

    `<w:embedSystemFonts/>` tells Word to bundle the machine's fonts into the
    file on save. Nothing embeds fonts here, so today it is only a latent
    instruction, and the day something acts on it the review copy starts
    shipping a third party's font binary to everyone it is sent to.
    """
    settings = "word/settings.xml"
    if settings not in parts:
        return 0
    root = parse_ooxml(parts[settings].decode("utf-8"), settings)
    removed = 0
    for tag in (Tag.EMBED_SYSTEM_FONTS, Tag.EMBED_TRUETYPE_FONTS):
        for element in children(root, tag):
            root.remove(element)
            removed += 1
    parts[settings] = serialize_ooxml(root).encode("utf-8")
    return removed


@dataclass(frozen=True)
class Restyling:
    """What the presentation pass changed, for the run to report.

    A dataclass rather than a dict because the fields are not one type: three
    counts and a list of the typefaces cleared. As a dict[str, int] the list
    field was a type error nobody could see.
    """

    tables_resized: int
    listings_highlighted: int
    font_directives_stripped: int
    theme_fonts_cleared: list[str]
    body_styles_remapped: int
    tall_math_exempted: int
    empty_comments_dropped: bool


def restyle_docx(path: Path) -> Restyling:
    """Apply the presentation pass to a finished .docx, in place."""
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
        infos = z.infolist()

    document = parse_ooxml(parts["word/document.xml"].decode("utf-8"), "word/document.xml")
    tables = resize_tables(document)
    listings = highlight_code_paragraphs(document)
    body_styles = map_styles_for_docs(document)
    exempted = exempt_tall_math_from_fixed_leading(document)
    parts["word/document.xml"] = serialize_ooxml(document).encode("utf-8")
    directives = strip_font_embedding(parts)
    # Two independent jobs: a document with no theme still has a font table.
    cleared = clear_theme_fonts(parts) + strip_unused_font_table(parts)
    dropped_comments = drop_empty_comments_part(parts)

    # `infos` is the zip listing as it was read, so it still names any part a
    # pass above deleted. Writing from `parts` and skipping what is gone keeps
    # the original entry order for everything that survives.
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            if info.filename in parts:
                z.writestr(info, parts[info.filename])
    return Restyling(
        tables_resized=tables,
        listings_highlighted=listings,
        font_directives_stripped=directives,
        theme_fonts_cleared=cleared,
        body_styles_remapped=body_styles,
        tall_math_exempted=exempted,
        empty_comments_dropped=dropped_comments,
    )


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------


def resolve_tool(name: str) -> str:
    """The absolute path to an external tool, or exit saying which is missing.

    Running by absolute path rather than by bare name means the binary cannot
    change under this process because of PATH ordering, which on a machine
    where the interactive and unattended PATHs disagree is a real difference
    rather than a theoretical one.
    """
    found = shutil.which(name)
    if found is None:
        sys.exit(f"Missing required tool: {name}")
    return found


def require_tools(engine: str) -> None:
    hints = {
        "pandoc": "brew install pandoc",
        engine: TEX_ENGINE_HINTS.get(engine, f"install {engine}"),
        "pdftocairo": "brew install poppler",
    }
    missing = [f"  {t}: {h}" for t, h in hints.items() if shutil.which(t) is None]
    if missing:
        sys.exit("Missing required tools:\n" + "\n".join(missing))


def convert(
    tex_path: Path,
    out_path: Path,
    bib: Path | None,
    dpi: int,
    workdir: Path,
    quiet: bool = False,
    engine: str | None = None,
) -> tuple[list[Check], list[str]]:
    """Run the full pipeline. Returns (checks, notes)."""
    notes: list[str] = []
    engine = engine or default_tex_engine()

    def say(msg: str) -> None:
        if not quiet:
            print(msg)

    src = tex_path.read_text(encoding="utf-8")
    facts = collect_source_facts(src)
    say(
        f"Source: {facts.tikz_figures} TikZ figure(s), {facts.graphics_figures} image figure(s), "
        f"{facts.top_level_tabulars} table(s), {len(facts.citation_keys)} citation key(s)"
    )

    figdir = workdir / "figs"
    figdir.mkdir(parents=True, exist_ok=True)

    figures: list[Figure] = []
    for i, block in enumerate(FIGURE_RE.findall(strip_comments(src))):
        tikz = TIKZ_RE.search(block)
        if not tikz:
            continue
        label_m = LABEL_RE.search(block)
        label = (
            re.sub(r"[^A-Za-z0-9]+", "-", label_m.group(1).split(":")[-1]) if label_m else f"fig{i}"
        )
        figures.append(Figure(index=i, label=label, block=block, tikz=tikz.group(0)))

    preamble = harvest_preamble(src)
    metrics = (
        measure_page_metrics(src, workdir, engine) if figures else PageMetrics(how="no figures")
    )
    if figures:
        say(
            f"Page metrics: columnwidth={metrics.column_width or 'not measured'}, "
            f"normalsize={metrics.normal_size or '?'}pt, "
            f"{len(metrics.size_steps)} size step(s) "
            f"({metrics.how}; engine: {engine})"
        )
        if metrics.column_width is None:
            notes.append(
                "Could not measure the page metrics, so figures were drawn at "
                "standalone's own width and font. Both the plot size and the "
                "text size relative to it will be wrong."
            )
        elif metrics.how != "after \\maketitle":
            notes.append(
                f"Page metrics were measured {metrics.how}. On a two-column "
                "class that reads \\textwidth instead of the column, which draws "
                "figures at double width; check them before shipping."
            )
        elif not metrics.size_steps:
            notes.append(
                "Measured \\columnwidth but not the font sizes, so figures use "
                "standalone's own ladder. If the paper's class is smaller, "
                "legends and tick labels will be too large for the plot."
            )

    failures = []
    for fig in figures:
        ok, detail = render_figure(fig, preamble, metrics, figdir, dpi, engine)
        say(f"  {fig.label}: {detail}" if ok else f"  {fig.label}: FAILED")
        if not ok:
            failures.append(detail)
        elif fig.fallback_preamble:
            notes.append(f"{fig.label} needed the fallback preamble; check it against the paper.")
    if failures:
        notes.extend(failures)

    rewritten, stats = rewrite_source(src, figures, figdir)
    if stats.flattened:
        notes.append(
            f"Flattened {stats.flattened} nested tabular cell(s); "
            "pandoc drops those tables otherwise."
        )
    if stats.unstarred:
        notes.append(
            f"Unstarred {stats.unstarred} table*/figure* environment(s); "
            "pandoc drops the caption of a starred float."
        )
    if stats.refs:
        notes.append(f"Resolved {stats.refs} table/figure/theorem cross-reference(s) to numbers.")
    if stats.listings:
        notes.append(
            f"Converted {stats.listings} alltt listing(s) to verbatim to keep "
            "their spacing; math escapes became literal characters."
        )
    if stats.unmapped:
        notes.append(
            "No character mapping for these macros in a listing, so they were dropped: "
            + ", ".join(sorted(stats.unmapped))
            + ". Add them to MATH_TO_UNICODE."
        )
    work_tex = workdir / "converted.tex"
    work_tex.write_text(rewritten, encoding="utf-8")

    cmd = [
        resolve_tool("pandoc"),
        str(work_tex),
        "-o",
        str(out_path),
        f"--resource-path={workdir}:{tex_path.parent}",
    ]
    reference = build_reference_doc(workdir)
    if reference:
        cmd.append(f"--reference-doc={reference}")
    else:
        notes.append("Could not build a reference document; pandoc's default styling was used.")
    if bib:
        cmd += [
            "--citeproc",
            f"--bibliography={bib}",
            "--metadata",
            "reference-section-title=References",
        ]
    proc = subprocess.run(cmd, capture_output=True, text=True)  # noqa: S603
    if proc.returncode != 0:
        notes.append(f"pandoc exited {proc.returncode}: {proc.stderr.strip()[:800]}")

    if out_path.exists():
        styling = restyle_docx(out_path)
        say(
            f"Styled: {styling.tables_resized} table(s) set to the full text width, "
            f"{styling.listings_highlighted} listing(s) highlighted"
        )
        if styling.body_styles_remapped:
            say(
                f"Styles: {styling.body_styles_remapped} paragraph(s) moved onto a style Google "
                f"Docs knows; {styling.tall_math_exempted} kept auto leading for tall maths"
            )
        if styling.font_directives_stripped or styling.theme_fonts_cleared:
            # A theme names a face per script, so spelling them all out buries
            # the rest of the report under ninety-odd names. The count is what
            # says the sweep was total; the first two are what a reader
            # recognises as the ones that were causing trouble.
            cleared = styling.theme_fonts_cleared
            distinct = list(dict.fromkeys(cleared))
            shown = ", ".join(distinct[:2])
            rest = f", and {len(distinct) - 2} more" if len(distinct) > 2 else ""
            say(
                f"Fonts: {styling.font_directives_stripped} embedding directive(s) removed, "
                f"{len(cleared)} typeface name(s) dropped from the theme and font table "
                f"({shown}{rest}) so the reader's template supplies them"
            )

    from .verification import verify

    checks = verify(out_path, facts, proc.stderr)
    return checks, notes


def report(checks: list[Check], notes: list[str]) -> int:
    print("\nVerification (run against the .docx itself):")
    width = max((len(c.name) for c in checks), default=0)
    for c in checks:
        print(f"  [{c.status:4}] {c.name:{width}}  {c.detail}")
    if notes:
        print("\nNotes:")
        for n in notes:
            print(f"  - {n}")
    failed = [c for c in checks if c.status == "FAIL"]
    if failed:
        print(f"\n{len(failed)} check(s) FAILED. The .docx is not trustworthy as written.")
        return 1
    print(
        "\nAll checks passed. Open it in Word, or upload it to Drive and use "
        "File > Save as Google Docs for a native copy."
    )
    return 0


# --------------------------------------------------------------------------
# Self-test — calibrate the checks before trusting a green run
# --------------------------------------------------------------------------


def self_test() -> int:
    """Confirm every check fires on input built to break it.

    Needs no pandoc and no TeX engine: `verify()` takes a path to a zip, a
    `SourceFacts`, and citeproc's stderr, so every case is two hand-written XML
    strings in a temporary archive. That is why this also runs in CI, where no
    TeX distribution exists.

    The cases live in `_calibration` rather than here, so this command and the
    pytest suite read one registry instead of drifting apart as two copies.
    """
    from ._calibration import run_calibration

    print("Self-test: every check must fail on input built to break it.\n")
    with tempfile.TemporaryDirectory() as td:
        ok, lines = run_calibration(Path(td), verbose=True)
    for line in lines:
        print(line)
    if ok:
        print("\nSelf-test passed: the checks fire on bad input and stay quiet on good input.")
        return 0
    print("\nSelf-test FAILED: a check did not behave as designed. Do not trust its verdict.")
    return 1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a LaTeX paper to a .docx for review in Word or Google Docs."
    )
    parser.add_argument("tex", nargs="?", type=Path, help="the .tex file to convert")
    parser.add_argument("-o", "--output", type=Path, help="output .docx (default: <tex>.docx)")
    parser.add_argument(
        "--bib",
        type=Path,
        help="BibTeX file (default: references.bib beside the source, if present)",
    )
    parser.add_argument(
        "--dpi", type=int, default=300, help="figure raster resolution (default: 300)"
    )
    parser.add_argument(
        "--tex-engine",
        choices=TEX_ENGINES,
        default=None,
        help=f"TeX engine for figures (default: first found of {', '.join(TEX_ENGINES)})",
    )
    parser.add_argument(
        "--work", type=Path, help="working directory, kept for inspection (default: <output>-work)"
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="verify that the checks detect the failures they claim to detect",
    )
    args = parser.parse_args()

    if args.self_test:
        return self_test()
    if not args.tex:
        parser.error("a .tex file is required (or use --self-test)")
    if not args.tex.exists():
        sys.exit(f"No such file: {args.tex}")

    engine = args.tex_engine or default_tex_engine()
    require_tools(engine)
    out = args.output or args.tex.with_suffix(".docx")
    bib = args.bib
    if bib is None:
        candidate = args.tex.parent / "references.bib"
        if candidate.exists():
            bib = candidate
            print(f"Using bibliography: {bib}")
    work = args.work or out.parent / f"{out.stem}-work"
    work.mkdir(parents=True, exist_ok=True)

    checks, notes = convert(args.tex, out, bib, args.dpi, work, engine=engine)
    print(f"\nWrote {out} ({out.stat().st_size // 1024} KB). Work files kept in {work}/")
    return report(checks, notes)


if __name__ == "__main__":
    sys.exit(main())
