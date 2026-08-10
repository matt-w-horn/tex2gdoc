"""Convert a real paper and compare it against a recorded fingerprint.

Skipped unless TEX2GDOC_REGRESSION_TEX names a paper and the tools are
installed, so it never runs in CI. The paper it targets lives in a separate
private repo and must not be copied here.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tex2gdoc.baseline import compare, measure_docx, regressed_checks
from tex2gdoc.tex2gdoc import convert

pytestmark = pytest.mark.regression


def test_paper_still_converts_to_the_recorded_shape(
    regression_tex: Path,
    regression_baseline: dict,
    tools_available: None,
    tmp_path: Path,
) -> None:
    bib = regression_tex.parent / "references.bib"
    checks, notes = convert(
        regression_tex,
        tmp_path / "out.docx",
        bib if bib.exists() else None,
        300,
        tmp_path / "work",
        quiet=True,
    )

    worse = regressed_checks(regression_baseline["checks"], {c.name: c.status for c in checks})
    assert not worse, "checks regressed:\n  " + "\n  ".join(worse)

    moved = compare(regression_baseline["counts"], measure_docx(tmp_path / "out.docx"))
    assert not moved, "output structure moved:\n  " + "\n  ".join(moved)
