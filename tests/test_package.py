"""Packaging checks: what has to be true of the installed distribution.

Behaviour is covered by the calibration suite and the table tests. What is left
to this file is the packaging itself, which fails silently rather than loudly:
an import that works from the source tree and not from a wheel, or a marker
file that is simply absent from the built artifact.

The docstring here used to say the converter had not landed yet and to replace
this file when it did. It landed; the file is now doing the job that is
actually left over.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import tex2gdoc


def test_package_imports_and_reports_a_version() -> None:
    assert isinstance(tex2gdoc.__version__, str)
    assert tex2gdoc.__version__


def test_py_typed_marker_ships_with_the_package() -> None:
    """Without the marker, every annotation in this package is invisible.

    A consumer running mypy against an installed tex2gdoc gets
    `module is installed, but missing library stubs or py.typed marker` and
    silently loses the `SafeElement` guarantee, which is the one thing the
    types here are load-bearing for. Nothing else notices it is gone: the
    tests pass, the tool runs, and only a downstream type check degrades.
    """
    spec = importlib.util.find_spec("tex2gdoc")
    assert spec is not None and spec.origin is not None
    assert (Path(spec.origin).parent / "py.typed").is_file()
