"""Check the .docx that ships against what the source claimed.

Separate from the converter on purpose. This is the layer that decides whether a
run can be trusted, and it reads the output file rather than any intermediate.

Every check here is calibrated: `tex2gdoc._calibration` builds input designed to
break each one and asserts that it reports FAIL, and a check added without a
calibration case fails the coverage gate. Silence is not evidence.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path

from .tex2gdoc import MARGIN, TEXT_WIDTH, SourceFacts, parse_ooxml, qn, wval


@dataclass
class Check:
    name: str
    status: str  # PASS | WARN | FAIL
    detail: str


def read_docx(path: Path) -> tuple[ET.Element, str, list[str], ET.Element | None]:
    """Return (document tree, its text, media file names, styles tree).

    Parsed rather than matched. Every detector below reads the tree, so a change
    in how pandoc spaces its tags cannot move a count, and a malformed document
    raises here instead of quietly matching nothing.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        document = parse_ooxml(
            z.read("word/document.xml").decode("utf-8", errors="replace"), "word/document.xml"
        )
        styles = (
            parse_ooxml(
                z.read("word/styles.xml").decode("utf-8", errors="replace"), "word/styles.xml"
            )
            if "word/styles.xml" in names
            else None
        )
        media = [n for n in names if n.startswith("word/media/")]
    return document, "".join(document.itertext()), media, styles


def verify(docx: Path, facts: SourceFacts, citeproc_stderr: str) -> list[Check]:
    """Check the output against what the source claimed. Silence is not evidence."""
    checks: list[Check] = []

    if not docx.exists() or docx.stat().st_size == 0:
        return [Check("output exists", "FAIL", "no .docx was written")]

    # Reading the output is now a parse, so a malformed part is caught here
    # rather than silently matching nothing further down. That is most of what
    # a schema validator was wanted for: this pass cannot write a document that
    # will not parse, and it says so out loud if one ever arrives.
    try:
        document, text, media, styles = read_docx(docx)
    except (ET.ParseError, ValueError, KeyError) as exc:
        return [Check("document parses", "FAIL", f"the .docx is not readable: {exc}")]

    flat = re.sub(r"\s+", " ", text)

    expected_images = facts.tikz_figures + facts.graphics_figures
    checks.append(
        Check(
            "figures embedded",
            "PASS" if len(media) >= expected_images else "FAIL",
            f"{len(media)} image(s) in the .docx, {expected_images} figure(s) in the source",
        )
    )

    tables = len(list(document.iter(qn("w:tbl"))))
    checks.append(
        Check(
            "tables present",
            "PASS" if tables >= facts.top_level_tabulars else "FAIL",
            f"{tables} table(s) in the .docx, {facts.top_level_tabulars} tabular(s) in the source",
        )
    )

    if tables:
        full = sum(
            1
            for e in document.iter(qn("w:tblW"))
            if e.get(qn("w:type")) == "dxa" and e.get(qn("w:w")) == str(TEXT_WIDTH)
        )
        cells = len(list(document.iter(qn("w:tc"))))
        sized = sum(1 for e in document.iter(qn("w:tcW")) if e.get(qn("w:type")) == "dxa")
        pct = sum(1 for e in document.iter() if e.get(qn("w:type")) == "pct")
        problems = []
        if full < tables:
            problems.append(f"{tables - full} table(s) not at the full text width")
        if sized < cells:
            problems.append(f"{cells - sized} cell(s) without an explicit width")
        if pct:
            problems.append(f"{pct} percentage width(s), which Google Docs renders unreliably")
        checks.append(
            Check(
                "table layout",
                "PASS" if not problems else "FAIL",
                f"{tables} table(s) at {TEXT_WIDTH} dxa, {cells} cell(s) sized in dxa"
                if not problems
                else "; ".join(problems),
            )
        )

        bad = [
            g
            for g in document.iter(qn("w:tblGrid"))
            if sum(int(c.get(qn("w:w"), "0")) for c in g.iter(qn("w:gridCol"))) != TEXT_WIDTH
        ]
        checks.append(
            Check(
                "column widths sum",
                "PASS" if not bad else "FAIL",
                f"every grid sums to {TEXT_WIDTH} dxa"
                if not bad
                else f"{len(bad)} grid(s) do not sum to the table width",
            )
        )

    if facts.tikz_figures or facts.graphics_figures:
        # Scoped to one <w:style> element each. The regex this replaced searched
        # DOTALL for a centred justification anywhere after a style id, so the
        # next centred style in the file satisfied it and an uncentred figure
        # style reported PASS.
        def style_is_centered(sid: str) -> bool:
            if styles is None:
                return False
            for style in styles.iter(qn("w:style")):
                if style.get(qn("w:styleId")) != sid:
                    continue
                ppr = style.find(qn("w:pPr"))
                return ppr is not None and wval(ppr.find(qn("w:jc"))) == "center"
            return False

        centered = all(style_is_centered(sid) for sid in ("CaptionedFigure", "ImageCaption"))
        checks.append(
            Check(
                "figures centered",
                "PASS" if centered else "FAIL",
                "figure and caption styles are centered"
                if centered
                else "figure or caption style is not centered — did the reference document build?",
            )
        )

    pgmar = next(document.iter(qn("w:pgMar")), None)
    margins_applied = pgmar is not None and pgmar.get(qn("w:left")) == str(MARGIN)
    checks.append(
        Check(
            "page geometry",
            "PASS" if margins_applied else "FAIL",
            f"margins set to {MARGIN} dxa"
            if margins_applied
            else "page margins were not applied — did the reference document build?",
        )
    )

    if facts.listing_probes:
        colored = sum(
            1
            for run in document.iter(qn("w:r"))
            if (rpr := run.find(qn("w:rPr"))) is not None
            and wval(rpr.find(qn("w:rStyle"))) == "VerbatimChar"
            and rpr.find(qn("w:color")) is not None
        )
        # FAIL, not WARN. report() exits nonzero only on FAIL, so a WARN here
        # meant a refactor could silently drop every coloured run and still
        # leave the command green.
        checks.append(
            Check(
                "code highlighted",
                "PASS" if colored else "FAIL",
                f"{colored} coloured code run(s)"
                if colored
                else "listings present but none of their runs are coloured",
            )
        )

    if facts.table_probes:
        lost = [p for p in facts.table_probes if p not in flat]
        checks.append(
            Check(
                "table contents survived",
                "PASS" if not lost else "FAIL",
                f"{len(facts.table_probes)} table(s) probed, all contents found"
                if not lost
                else f"content missing from {len(lost)} table(s); "
                f"absent word(s): {', '.join(lost[:3])}"
                " — usually an unsupported construct in a cell, such as a nested tabular",
            )
        )

    if facts.caption_probes:
        lost = [p for p in facts.caption_probes if p not in flat]
        checks.append(
            Check(
                "captions survived",
                "PASS" if not lost else "FAIL",
                f"{len(facts.caption_probes)} caption(s) probed, all found"
                if not lost
                else f"{len(lost)} caption(s) missing, first: {lost[0]!r}",
            )
        )

    if facts.listing_probes:
        lost = [p for p in facts.listing_probes if p not in flat]
        checks.append(
            Check(
                "code listings readable",
                "PASS" if not lost else "FAIL",
                f"{len(facts.listing_probes)} listing(s) probed, spacing intact"
                if not lost
                else f"{len(lost)} listing(s) lost their spacing, first: {lost[0]!r}",
            )
        )

    dangling = sorted(set(re.findall(r"\[(?:tab|fig|thm|sec|app|eq):[A-Za-z0-9_\-]+\]", text)))
    checks.append(
        Check(
            "cross-references resolved",
            "PASS" if not dangling else "FAIL",
            "no raw label keys in the text"
            if not dangling
            else f"{len(dangling)} unresolved: {', '.join(dangling[:4])}",
        )
    )

    missing = re.findall(r"citation ([^\s]+) not found", citeproc_stderr)
    if facts.citation_keys:
        checks.append(
            Check(
                "citations resolved",
                "PASS" if not missing else "FAIL",
                f"{len(facts.citation_keys)} key(s) cited"
                + (
                    f", unresolved: {', '.join(sorted(set(missing))[:5])}"
                    if missing
                    else ", all resolved"
                ),
            )
        )
        # Stripping XML tags joins a heading to whatever precedes it, so a
        # word-boundary match fails on text that is plainly there. A heading
        # alone proves nothing either: the list under it has to have body.
        pos = max(text.rfind("References"), text.rfind("Bibliography"))
        tail = len(text) - pos if pos >= 0 else 0
        # One rendered entry runs well past 40 characters, so scale the floor to
        # the number of keys rather than picking a constant that only suits a
        # long paper.
        floor = 40 * len(facts.citation_keys)
        checks.append(
            Check(
                "bibliography rendered",
                "PASS" if pos >= 0 and tail >= floor else "FAIL",
                f"reference section found, {tail} chars of entries (floor {floor})"
                if pos >= 0
                else "no reference section — was a bibliography passed?",
            )
        )

    math = len(list(document.iter(qn("m:oMath"))))
    if facts.inline_math:
        ratio = math / facts.inline_math
        status = "FAIL" if math == 0 else ("WARN" if ratio < 0.5 else "PASS")
        checks.append(
            Check(
                "math converted",
                status,
                f"{math} Word equation(s), {facts.inline_math} inline-math span(s) in the source "
                f"(ratio {ratio:.2f}; exact parity is not expected)",
            )
        )

    headings = sum(1 for e in document.iter(qn("w:pStyle")) if wval(e) == "Heading1")
    if facts.sections:
        # Also FAIL rather than WARN: fewer top-level headings than the source
        # has sections means the outline a reviewer navigates by is incomplete,
        # and that is not a warning.
        checks.append(
            Check(
                "sections present",
                "PASS" if headings >= facts.sections else "FAIL",
                f"{headings} top-level heading(s), {facts.sections} \\section(s) in the source",
            )
        )

    ratio = len(text) / facts.body_chars if facts.body_chars else 0
    checks.append(
        Check(
            "body text volume",
            "PASS" if ratio >= 0.5 else "FAIL",
            f"{len(text)} chars extracted, {facts.body_chars} in the source (ratio {ratio:.2f})",
        )
    )

    return checks
