"""Fixtures for the tool-dependent tests.

Everything here skips rather than fails when the world is not set up: the
regression target is a paper in a different, private repo that CI cannot see,
and CI has no TeX distribution.
"""

from __future__ import annotations

import json
import os
import shutil
from pathlib import Path
from typing import Any

import pytest


@pytest.fixture(scope="session")
def regression_tex() -> Path:
    """The paper to convert, from TEX2GDOC_REGRESSION_TEX."""
    raw = os.environ.get("TEX2GDOC_REGRESSION_TEX")
    if not raw:
        pytest.skip("set TEX2GDOC_REGRESSION_TEX to a .tex file to run the regression test")
    path = Path(raw).expanduser()
    if not path.is_file():
        pytest.skip(f"TEX2GDOC_REGRESSION_TEX points at {path}, which is not a file")
    return path


@pytest.fixture(scope="session")
def regression_baseline(regression_tex: Path) -> dict[str, Any]:
    """The recorded fingerprint, which lives beside the paper, not in this repo."""
    path = regression_tex.parent / ".verify" / "tex2gdoc-baseline.json"
    if not path.exists():
        pytest.skip(f"no baseline at {path}; run scripts/record_baseline.py once")
    baseline: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    return baseline


@pytest.fixture(scope="session")
def tools_available() -> None:
    """Skip unless pandoc, a TeX engine and pdftocairo are all on PATH."""
    missing = [t for t in ("pandoc", "pdftocairo") if shutil.which(t) is None]
    if not any(shutil.which(e) for e in ("tectonic", "pdflatex")):
        missing.append("tectonic or pdflatex")
    if missing:
        pytest.skip("not installed: " + ", ".join(missing))
