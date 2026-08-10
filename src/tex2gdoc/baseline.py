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
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from .tex2gdoc import STYLE_PATCHES, TEXT_WIDTH, parse_ooxml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

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


def _pstyle_count(root: ET.Element, style_id: str) -> int:
    return sum(1 for e in root.iter(f"{{{W}}}pStyle") if e.get(f"{{{W}}}val") == style_id)


def _colored_code_runs(root: ET.Element) -> int:
    """Runs carrying the VerbatimChar character style AND an explicit colour.

    The tree equivalent of the old adjacency regex, and immune to the tag spacing
    that regex depended on.
    """
    total = 0
    for run in root.iter(f"{{{W}}}r"):
        rpr = run.find(f"{{{W}}}rPr")
        if rpr is None:
            continue
        style = rpr.find(f"{{{W}}}rStyle")
        if style is None or style.get(f"{{{W}}}val") != "VerbatimChar":
            continue
        if rpr.find(f"{{{W}}}color") is not None:
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

    text = re.sub(r"<[^>]+>", "", raw_document)
    ref_pos = max(text.rfind("References"), text.rfind("Bibliography"))

    grid_sums = [
        sum(int(c.get(f"{{{W}}}w", "0")) for c in grid.iter(f"{{{W}}}gridCol"))
        for grid in doc.iter(f"{{{W}}}tblGrid")
    ]

    return {
        "docx_size_bytes": path.stat().st_size,
        "media_files": sum(1 for n in names if n.startswith("word/media/")),
        "drawings": sum(1 for _ in doc.iter(f"{{{W}}}drawing")),
        "omath": sum(1 for _ in doc.iter(f"{{{M}}}oMath")),
        "tbl": sum(1 for _ in doc.iter(f"{{{W}}}tbl")),
        "tc": sum(1 for _ in doc.iter(f"{{{W}}}tc")),
        "tc_with_tcW_dxa": sum(
            1 for e in doc.iter(f"{{{W}}}tcW") if e.get(f"{{{W}}}type") == "dxa"
        ),
        "tblW_at_text_width": sum(
            1
            for e in doc.iter(f"{{{W}}}tblW")
            if e.get(f"{{{W}}}type") == "dxa" and e.get(f"{{{W}}}w") == str(TEXT_WIDTH)
        ),
        "tblgrids": len(grid_sums),
        "tblgrids_summing_to_text_width": sum(1 for s in grid_sums if s == TEXT_WIDTH),
        "pct_widths": sum(1 for e in doc.iter() if e.get(f"{{{W}}}type") == "pct"),
        "gridspans": sum(1 for _ in doc.iter(f"{{{W}}}gridSpan")),
        "heading1": _pstyle_count(doc, "Heading1"),
        "heading2": _pstyle_count(doc, "Heading2"),
        "sourcecode_paragraphs": _pstyle_count(doc, "SourceCode"),
        "colored_code_runs": _colored_code_runs(doc),
        "document_xml_chars": len(raw_document),
        "stripped_text_chars": len(text),
        "dangling_labels": len(set(DANGLING_LABEL_RE.findall(text))),
        "reference_tail_chars": (len(text) - ref_pos) if ref_pos >= 0 else 0,
        "styles_defined": (
            sum(1 for e in styles.iter(f"{{{W}}}style") if e.get(f"{{{W}}}styleId"))
            if styles is not None
            else 0
        ),
        "patched_styles_present": (
            sum(
                1 for e in styles.iter(f"{{{W}}}style") if e.get(f"{{{W}}}styleId") in STYLE_PATCHES
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
