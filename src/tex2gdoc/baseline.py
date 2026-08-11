"""Measure a .docx into a structural fingerprint, and compare two fingerprints.

The regression target for this converter is a paper in another, private repo, so
the .docx it produces can never be committed here and cannot be compared
byte-for-byte anyway: pandoc, tectonic and citeproc versions all move, and zip
timestamps differ between runs. What is stable is the *shape* of the output, and
that is what this module records.

Structure is counted by walking the XML tree rather than by matching serialized
text, so a change in tag spacing or attribute order cannot move a number. One
measurement deliberately breaks that rule: `stripped_text_chars` uses the same
crude tag strip `verify()` uses, so it stays comparable with what `verify()`
itself sees.

A `colored_code_runs_adjacent` metric lived here through the ElementTree
migration, reproducing the old adjacency regex so the tree walk could be checked
against the thing it replaced. Both returned 434 on the real paper, so it has
been removed rather than left as a second way to count the same thing.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from .ooxml import Attr, SafeElement, Tag, Value, attribute, child, descendants, value_of
from .tex2gdoc import STYLE_PATCHES, TEXT_WIDTH, StyleId, parse_ooxml

# How far each metric may move before it counts as a regression. "exact" is for
# invariants of correctness; a percentage is for numbers that a pandoc or
# citeproc version bump may legitimately perturb without anything being wrong.
TOLERANCES: dict[str, float | str] = {
    "media_files": "exact",
    "drawings": "exact",
    "tbl": "exact",
    "tc": "exact",
    "tc_with_tcW_dxa": "exact",
    "tblW_at_text_width": "exact",
    "tblgrids_summing_to_text_width": "exact",
    "tblgrids": "exact",
    "pct_widths": "exact",
    "gridspans": "exact",
    "heading1": "exact",
    "heading2": "exact",
    "sourcecode_paragraphs": "exact",
    "dangling_labels": "exact",
    "omath": 0.05,
    "colored_code_runs": 0.05,
    "stripped_text_chars": 0.10,
    "reference_tail_chars": 0.15,
    "patched_styles_present": "exact",
    # Informational only. `styles_defined` counts whatever pandoc's bundled
    # reference document happens to define, so it tracks the pandoc version
    # rather than anything about this converter: measured 69 under the pandoc
    # that produced the first shipped artifact and 81 under 3.10.1, with both
    # counting methods agreeing on each file. `patched_styles_present` is the
    # part that is actually ours, so that is what gets held exact.
    "styles_defined": "info",
    "docx_size_bytes": "info",
    "document_xml_chars": "info",
}

DANGLING_LABEL_RE = re.compile(r"\[(?:tab|fig|thm|sec|app|eq):[A-Za-z0-9_\-]+\]")


def _pstyle_count(root: SafeElement, style_id: str) -> int:
    return sum(1 for e in descendants(root, Tag.PARAGRAPH_STYLE) if value_of(e) == style_id)


def _colored_code_runs(root: SafeElement) -> int:
    """Runs carrying the VerbatimChar character style AND an explicit colour.

    The tree equivalent of the old adjacency regex, and immune to the tag spacing
    that regex depended on.
    """
    total = 0
    for run in descendants(root, Tag.RUN):
        rpr = child(run, Tag.RUN_PROPERTIES)
        if rpr is None:
            continue
        if value_of(child(rpr, Tag.RUN_STYLE)) != StyleId.VERBATIM_CHAR:
            continue
        if child(rpr, Tag.COLOR) is not None:
            total += 1
    return total


def measure_docx(path: Path) -> dict[str, int]:
    """Count the structural features that must survive a refactor."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        raw_document = z.read("word/document.xml").decode("utf-8", errors="replace")
        raw_styles = (
            z.read("word/styles.xml").decode("utf-8", errors="replace")
            if "word/styles.xml" in names
            else ""
        )

    doc = parse_ooxml(raw_document, "word/document.xml")
    styles = parse_ooxml(raw_styles, "word/styles.xml") if raw_styles else None

    # Tree text, the same extraction verify() uses, so the two stay comparable.
    # This was a regex tag strip until 2026-08-10; on the real paper the two
    # differ by 0.18%, well inside this metric's 10% tolerance, so recorded
    # baselines survived the change.
    text = "".join(t for t in doc.itertext() if isinstance(t, str))
    ref_pos = max(text.rfind("References"), text.rfind("Bibliography"))

    grid_sums = [
        sum(int(attribute(c, Attr.WIDTH) or 0) for c in descendants(grid, Tag.GRID_COLUMN))
        for grid in descendants(doc, Tag.TABLE_GRID)
    ]

    return {
        "docx_size_bytes": path.stat().st_size,
        "media_files": sum(1 for n in names if n.startswith("word/media/")),
        "drawings": len(descendants(doc, Tag.DRAWING)),
        "omath": len(descendants(doc, Tag.MATH)),
        "tbl": len(descendants(doc, Tag.TABLE)),
        "tc": len(descendants(doc, Tag.TABLE_CELL)),
        "tc_with_tcW_dxa": sum(
            1 for e in descendants(doc, Tag.CELL_WIDTH) if attribute(e, Attr.TYPE) == Value.DXA
        ),
        "tblW_at_text_width": sum(
            1
            for e in descendants(doc, Tag.TABLE_WIDTH)
            if attribute(e, Attr.TYPE) == Value.DXA and attribute(e, Attr.WIDTH) == str(TEXT_WIDTH)
        ),
        "tblgrids": len(grid_sums),
        "tblgrids_summing_to_text_width": sum(1 for s in grid_sums if s == TEXT_WIDTH),
        "pct_widths": sum(
            1
            for e in (SafeElement(x) for x in doc.iter())
            if attribute(e, Attr.TYPE) == Value.PERCENT
        ),
        "gridspans": len(descendants(doc, Tag.GRID_SPAN)),
        "heading1": _pstyle_count(doc, "Heading1"),
        "heading2": _pstyle_count(doc, "Heading2"),
        "sourcecode_paragraphs": _pstyle_count(doc, "SourceCode"),
        "colored_code_runs": _colored_code_runs(doc),
        "document_xml_chars": len(raw_document),
        "stripped_text_chars": len(text),
        "dangling_labels": len(set(DANGLING_LABEL_RE.findall(text))),
        "reference_tail_chars": (len(text) - ref_pos) if ref_pos >= 0 else 0,
        "styles_defined": (
            sum(1 for e in descendants(styles, Tag.STYLE) if attribute(e, Attr.STYLE_ID))
            if styles is not None
            else 0
        ),
        "patched_styles_present": (
            sum(
                1
                for e in descendants(styles, Tag.STYLE)
                if attribute(e, Attr.STYLE_ID) in STYLE_PATCHES
            )
            if styles is not None
            else 0
        ),
    }


def compare(baseline: dict[str, int], current: dict[str, int]) -> list[str]:
    """Return one message per metric that moved outside its tolerance."""
    problems: list[str] = []
    for name, expected in baseline.items():
        tol = TOLERANCES.get(name, "exact")
        if tol == "info":
            continue
        if name not in current:
            problems.append(f"{name}: not measured in the current run (baseline {expected})")
            continue
        got = current[name]
        if tol == "exact":
            if got != expected:
                problems.append(f"{name}: {got}, baseline {expected} (must match exactly)")
        else:
            slack = abs(expected) * float(tol)
            if abs(got - expected) > slack:
                problems.append(
                    f"{name}: {got}, baseline {expected} (outside +/-{float(tol) * 100:.0f}%)"
                )
    return problems


STATUS_RANK = {"PASS": 0, "WARN": 1, "FAIL": 2}


def regressed_checks(baseline: dict[str, str], current: dict[str, str]) -> list[str]:
    """Return one message per check that got worse, or stopped being emitted.

    A check that vanishes is treated as a regression, not as a pass: the failure
    mode this guards against is a source-parsing bug that empties a `facts` field
    so the check is never appended at all.
    """
    problems: list[str] = []
    for name, was in baseline.items():
        now = current.get(name)
        if now is None:
            problems.append(f"{name}: no longer emitted (baseline {was})")
        elif STATUS_RANK.get(now, 3) > STATUS_RANK.get(was, 3):
            problems.append(f"{name}: {now}, baseline {was}")
    return problems
