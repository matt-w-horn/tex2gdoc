"""The checks are only worth their verdict if they can be shown to fail.

These run with no pandoc and no TeX engine, so they are the layer CI actually
executes.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tex2gdoc._calibration import (
    CALIBRATIONS,
    EARLY_RETURN_CHECKS,
    declared_check_names,
    good_document_xml,
    good_facts,
    good_styles_xml,
    write_docx,
)
from tex2gdoc.verification import verify


def test_every_check_has_a_calibration_case() -> None:
    """A new check with no way to prove it can fail must not reach main."""
    declared = declared_check_names()
    registered = set(CALIBRATIONS)
    assert not declared - registered, (
        f"verify() can emit these with no calibration case: {sorted(declared - registered)}"
    )
    assert not registered - declared, (
        f"calibrated but verify() no longer emits: {sorted(registered - declared)}"
    )


def test_known_good_input_emits_exactly_the_expected_checks(tmp_path: Path) -> None:
    """Not merely "no FAIL": the exact set of names.

    Several checks are conditional on a SourceFacts field being non-empty. A
    parsing bug that empties one does not make its check fail, it makes the
    check disappear, and a suite that only looks for FAIL calls that clean.
    """
    document = good_document_xml()
    docx = write_docx(tmp_path, document, good_styles_xml())
    checks = verify(docx, good_facts(document), "")

    # The early-return checks report a file that is missing or unreadable, so
    # they can never appear alongside the rest.
    assert {c.name for c in checks} == declared_check_names() - EARLY_RETURN_CHECKS
    assert [f"{c.name}={c.status}" for c in checks if c.status != "PASS"] == []


@pytest.mark.parametrize("name", sorted(CALIBRATIONS))
def test_check_fires_on_input_built_to_break_it(name: str, tmp_path: Path) -> None:
    calibration = CALIBRATIONS[name]
    docx, facts, citeproc_stderr = calibration.build(tmp_path)
    emitted = verify(docx, facts, citeproc_stderr)
    got = next((c for c in emitted if c.name == name), None)
    assert got is not None, (
        f"{name!r} was not emitted at all for input built to break it; "
        f"got {sorted(c.name for c in emitted)}"
    )
    assert got.status == calibration.worst, (
        f"{name!r} reported {got.status} ({got.detail}); "
        f"expected {calibration.worst} for: {calibration.why}"
    )
