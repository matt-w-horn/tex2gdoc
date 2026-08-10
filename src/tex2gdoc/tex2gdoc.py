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
parser rather than quietly matching nothing.

The verification step is the point, and it runs against the output file rather
than a copy or an intermediate. All 18 checks are calibrated: `--self-test`
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
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - annotations only
    # Imported lazily at runtime inside convert(). `verification` reads this
    # module's constants, so a module-level import here would be a cycle.
    from .verification import Check

# --------------------------------------------------------------------------
# OOXML parsing
# --------------------------------------------------------------------------

# Every prefix pandoc declares, registered so a round-trip writes them back
# unchanged rather than inventing ns0, ns1. Checked against pandoc 3.10.1's
# output: nine prefixes on document.xml, two on styles.xml, and no
# `mc:Ignorable`, which matters because that attribute names prefixes as text
# and would break if one were renamed.
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>\n'


def qn(prefixed: str) -> str:
    """Turn `w:tbl` into ElementTree's Clark notation.

    One function rather than a `w()` and an `m()`, because this module already
    uses both `w` and `m` as local names for regex matches and cell widths, and
    a shadowed helper fails in a way that reads like a typo.
    """
    prefix, _, local = prefixed.partition(":")
    return f"{{{NS[prefix]}}}{local}"


def wval(element: ET.Element | None) -> str | None:
    """The `w:val` of an element, or None if the element is absent."""
    return element.get(qn("w:val")) if element is not None else None


def parse_fragment(fragment: str) -> list[ET.Element]:
    """Parse a bare run of XML elements that share the w namespace.

    The style patches are written as literal XML because that is the clearest
    way to state a declarative bundle of properties. Parsing them here means the
    document is still assembled as a tree, so nothing downstream depends on how
    they were spelled.
    """
    wrapper = f'<wrap xmlns:w="{NS["w"]}">{fragment}</wrap>'
    return list(ET.fromstring(wrapper))


def serialize_ooxml(root: ET.Element) -> str:
    return XML_DECLARATION + ET.tostring(root, encoding="unicode")


def parse_ooxml(raw: str, part: str) -> ET.Element:
    """Parse an OOXML part, refusing anything that carries a DTD.

    Both external-entity (XXE) and entity-expansion ("billion laughs") attacks
    need a DOCTYPE, and no part pandoc writes has one, so rejecting a DTD removes
    the whole class without pulling in `defusedxml`. A DOCTYPE is only valid in
    the prolog, so scanning the head is sufficient: anything later is malformed
    and the parser rejects it anyway.

    A .docx is normally produced locally by this script, but `verify()` reads
    whatever path it is given, including a file that came back from a co-author,
    so provenance is not assumed.
    """
    if re.search(r"<!DOCTYPE", raw[:8192], re.IGNORECASE):
        raise ValueError(f"{part} carries a DTD; refusing to parse it")
    return ET.fromstring(raw)


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
    return max(pool, key=len)


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


def run_tex(engine: str, stem: str, workdir: Path) -> subprocess.CompletedProcess:
    """Compile <stem>.tex inside workdir with the chosen engine.

    tectonic swallows the engine's own chatter unless asked for it, so `--print`
    is what lets a \\typeout reach stdout, and `--keep-logs` leaves behind the
    .log that a failure needs to quote.
    """
    if engine == "tectonic":
        cmd = ["tectonic", "--print", "--keep-logs", f"{stem}.tex"]
    else:
        cmd = ["pdflatex", "-interaction=nonstopmode", "-halt-on-error", f"{stem}.tex"]
    return subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)


def tex_output(result: subprocess.CompletedProcess, stem: str, workdir: Path) -> str:
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

    def build(pre: str) -> subprocess.CompletedProcess:
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

    subprocess.run(
        ["pdftocairo", "-png", "-r", str(dpi), "-singlefile", f"{stem}.pdf", stem],
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

    def convert_math(m: re.Match) -> str:
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

    def convert_block(m: re.Match) -> str:
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

    def sub(m: re.Match) -> str:
        nonlocal replaced
        key = m.group(1)
        if key not in numbering:
            return m.group(0)
        replaced += 1
        return str(numbering[key])

    return REF_RE.sub(sub, src), replaced


def rewrite_source(src: str, figures: list[Figure], figdir: Path) -> tuple[str, dict[str, object]]:
    """Apply every transformation pandoc needs, and report what each one did."""
    by_block = {f.block: f for f in figures if f.png}

    def replace(match: re.Match) -> str:
        fig = by_block.get(match.group(0))
        if fig is None:
            return match.group(0)
        include = f"\\includegraphics[width=\\linewidth]{{{figdir.name}/{fig.png.name}}}"
        return TIKZ_RE.sub(lambda _: include, match.group(0), count=1)

    src = FIGURE_RE.sub(replace, src)
    src, flattened = flatten_nested_tabulars(src)
    src, unstarred = unstar_float_environments(src)
    src, refs = resolve_float_references(src)
    src, listings, unmapped = alltt_to_verbatim(src)
    return src, {
        "flattened": flattened,
        "unstarred": unstarred,
        "refs": refs,
        "listings": listings,
        "unmapped": unmapped,
    }


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

SECT_PR = (
    f'<w:sectPr><w:pgSz w:w="{PAGE_WIDTH}" w:h="{PAGE_HEIGHT}" />'
    f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}"'
    f' w:header="360" w:footer="360" w:gutter="0" /></w:sectPr>'
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
def _body_spacing(after: int, line: int = 276) -> str:
    """Paragraph spacing in twentieths of a point. line=276 is 1.15 lines."""
    return (
        f'<w:pPr><w:spacing w:before="0" w:after="{after}"'
        f' w:line="{line}" w:lineRule="auto" /></w:pPr>'
    )


STYLE_PATCHES: dict[str, str] = {
    # 1.15 line and 8pt after a paragraph are Google Docs' own body defaults, so
    # the import matches a document created there rather than announcing itself.
    "Normal": _body_spacing(after=160),
    "BodyText": _body_spacing(after=160),
    # List items stay tight: 8pt between consecutive bullets reads as broken.
    "Compact": _body_spacing(after=0),
    # Centring a figure is layout, not decoration, and `verify()` checks for it.
    "CaptionedFigure": '<w:pPr><w:keepNext /><w:jc w:val="center" /></w:pPr>',
    "Figure": '<w:pPr><w:keepNext /><w:jc w:val="center" /></w:pPr>',
    "ImageCaption": '<w:pPr><w:jc w:val="center" /></w:pPr>',
    # Code has to be monospace to be readable at all, and the shaded block is
    # what separates a listing from prose.
    "VerbatimChar": f'<w:rPr><w:rFonts w:ascii="{CODE_FONT}" w:hAnsi="{CODE_FONT}"'
    f' w:cs="{CODE_FONT}" /></w:rPr>',
    "SourceCode": "<w:pPr><w:keepLines />"
    '<w:spacing w:before="80" w:after="80" w:line="240" w:lineRule="auto" />'
    '<w:shd w:val="clear" w:color="auto" w:fill="F7F8FA" /></w:pPr>'
    f'<w:rPr><w:rFonts w:ascii="{CODE_FONT}" w:hAnsi="{CODE_FONT}"'
    f' w:cs="{CODE_FONT}" /></w:rPr>',
}


def patch_styles(styles_xml: str) -> str:
    """Replace the paragraph and run properties of the styles listed above.

    Nothing else in styles.xml is touched. The document default font size, every
    heading definition and pandoc's own table style are left exactly as pandoc
    wrote them, because those are what Google Docs reads to decide what its own
    named styles mean.
    """
    root = parse_ooxml(styles_xml, "word/styles.xml")
    for style in root.iter(qn("w:style")):
        patch = STYLE_PATCHES.get(style.get(qn("w:styleId")))
        if patch is None:
            continue
        for tag in ("w:pPr", "w:rPr"):
            for existing in style.findall(qn(tag)):
                style.remove(existing)
        style.extend(parse_fragment(patch))
    return serialize_ooxml(root)


def build_reference_doc(workdir: Path) -> Path | None:
    """Write a pandoc reference document carrying the page geometry and styles.

    Derived from pandoc's own default rather than kept as a stored file, so it
    always matches the installed pandoc instead of drifting from it.
    """
    default = subprocess.run(
        ["pandoc", "--print-default-data-file", "reference.docx"], capture_output=True
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
                body = root.find(qn("w:body"))
                if body is not None:
                    for existing in body.findall(qn("w:sectPr")):
                        body.remove(existing)
                    body.extend(parse_fragment(SECT_PR))
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
CODE_COLORS = {
    "keyword": "0033B3",
    "type": "0F7B6C",
    "number": "B26B00",
    "symbol": "7A3E9D",
    "comment": "6A737D",
    "plain": "1A1A1A",
}
TOKEN_RE = re.compile(r"--[^\n]*|[A-Za-zℝℕℤℚ_][A-Za-z0-9ℝℕℤℚ_.']*|\d+|\s+|.")


def classify_token(tok: str) -> str:
    if tok.startswith("--"):
        return "comment"
    if tok in LEAN_KEYWORDS:
        return "keyword"
    if tok in LEAN_TYPES or (tok[:1].isupper() and tok[:1].isalpha()):
        return "type"
    if tok.isdigit():
        return "number"
    if not tok.strip():
        return "plain"
    if not (tok[0].isalnum() or tok[0] == "_"):
        return "symbol"
    return "plain"


def make_code_run(token: str) -> ET.Element:
    """One coloured run for one token."""
    kind = classify_token(token)
    run = ET.Element(qn("w:r"))
    rpr = ET.SubElement(run, qn("w:rPr"))
    ET.SubElement(rpr, qn("w:rStyle"), {qn("w:val"): "VerbatimChar"})
    ET.SubElement(rpr, qn("w:color"), {qn("w:val"): CODE_COLORS[kind]})
    if kind == "keyword":
        ET.SubElement(rpr, qn("w:b"))
    elif kind == "comment":
        ET.SubElement(rpr, qn("w:i"))
    text = ET.SubElement(run, qn("w:t"), {XML_SPACE: "preserve"})
    text.text = token
    return run


def highlight_code_paragraphs(root: ET.Element) -> int:
    """Recolour the runs of every `SourceCode` paragraph, token by token.

    A run carrying no `w:t` is passed through untouched. Pandoc puts the line
    break between listing lines in a run of its own, so dropping those would run
    the whole listing onto one line.
    """
    count = 0
    for para in root.iter(qn("w:p")):
        ppr = para.find(qn("w:pPr"))
        if ppr is None or wval(ppr.find(qn("w:pStyle"))) != "SourceCode":
            continue
        count += 1
        rebuilt: list[ET.Element] = []
        for child in list(para):
            node = child.find(qn("w:t")) if child.tag == qn("w:r") else None
            if node is None:
                rebuilt.append(child)
                continue
            rebuilt.extend(make_code_run(tok) for tok in TOKEN_RE.findall(node.text or ""))
        para[:] = rebuilt
    return count


def cell_text(cell: ET.Element) -> str:
    return re.sub(r"\s+", " ", "".join(cell.itertext())).strip()


def cell_is_monospace(cell: ET.Element) -> bool:
    return any(wval(s) == "VerbatimChar" for s in cell.iter(qn("w:rStyle")))


def column_widths(rows: list[list[ET.Element]], columns: int) -> list[int]:
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
    parent: ET.Element, tag: str, attrib: dict[str, str], first: bool = False
) -> None:
    """Drop every existing `tag` child of parent and add one with these attributes."""
    for existing in parent.findall(tag):
        parent.remove(existing)
    element = ET.Element(tag, attrib)
    if first:
        parent.insert(0, element)
    else:
        parent.append(element)


def resize_tables(root: ET.Element) -> int:
    """Make every table span the text width, with columns sized to their content."""
    count = 0
    for tbl in root.iter(qn("w:tbl")):
        rows = [tr.findall(qn("w:tc")) for tr in tbl.findall(qn("w:tr"))]
        if not rows:
            continue
        widths = column_widths(rows, max(len(r) for r in rows))
        count += 1

        grid = tbl.find(qn("w:tblGrid"))
        if grid is not None:
            tail = grid.tail
            grid.clear()
            grid.tail = tail
            for width in widths:
                ET.SubElement(grid, qn("w:gridCol"), {qn("w:w"): str(width)})

        tblpr = tbl.find(qn("w:tblPr"))
        if tblpr is None:
            tblpr = ET.Element(qn("w:tblPr"))
            tbl.insert(0, tblpr)
        _replace_child(tblpr, qn("w:tblW"), {qn("w:type"): "dxa", qn("w:w"): str(TEXT_WIDTH)})
        _replace_child(tblpr, qn("w:tblLayout"), {qn("w:type"): "fixed"})

        # Every cell carries its own width as well as the grid: Google Docs
        # honours the cell width, and a table with only a grid comes in squeezed.
        for tr in tbl.findall(qn("w:tr")):
            index = 0
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                if tcpr is None:
                    tcpr = ET.Element(qn("w:tcPr"))
                    tc.insert(0, tcpr)
                span = int(wval(tcpr.find(qn("w:gridSpan"))) or 1)
                width = sum(widths[index : index + span]) or widths[-1]
                index += span
                _replace_child(
                    tcpr, qn("w:tcW"), {qn("w:type"): "dxa", qn("w:w"): str(width)}, first=True
                )
    return count


def restyle_docx(path: Path) -> dict[str, int]:
    """Apply the presentation pass to a finished .docx, in place."""
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
        infos = z.infolist()

    document = parse_ooxml(parts["word/document.xml"].decode("utf-8"), "word/document.xml")
    tables = resize_tables(document)
    listings = highlight_code_paragraphs(document)
    parts["word/document.xml"] = serialize_ooxml(document).encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, parts[info.filename])
    return {"tables_resized": tables, "listings_highlighted": listings}


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------


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
    if stats["flattened"]:
        notes.append(
            f"Flattened {stats['flattened']} nested tabular cell(s); "
            "pandoc drops those tables otherwise."
        )
    if stats["unstarred"]:
        notes.append(
            f"Unstarred {stats['unstarred']} table*/figure* environment(s); "
            "pandoc drops the caption of a starred float."
        )
    if stats["refs"]:
        notes.append(
            f"Resolved {stats['refs']} table/figure/theorem cross-reference(s) to numbers."
        )
    if stats["listings"]:
        notes.append(
            f"Converted {stats['listings']} alltt listing(s) to verbatim to keep "
            "their spacing; math escapes became literal characters."
        )
    if stats["unmapped"]:
        notes.append(
            "No character mapping for these macros in a listing, so they were dropped: "
            + ", ".join(sorted(stats["unmapped"]))
            + ". Add them to MATH_TO_UNICODE."
        )
    work_tex = workdir / "converted.tex"
    work_tex.write_text(rewritten, encoding="utf-8")

    cmd = [
        "pandoc",
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
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        notes.append(f"pandoc exited {proc.returncode}: {proc.stderr.strip()[:800]}")

    if out_path.exists():
        styling = restyle_docx(out_path)
        say(
            f"Styled: {styling['tables_resized']} table(s) set to the full text width, "
            f"{styling['listings_highlighted']} listing(s) highlighted"
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
