"""Scaffold smoke test.

This asserts only that the package installs and imports. That is a real check
of the build and of CI's install step, and nothing more — it says nothing about
conversion behaviour, because there is none yet. It also keeps `pytest` from
exiting 5 ("no tests collected"), which would fail CI on an empty suite.

Replace this with real tests when the converter lands.
"""

import tex2gdoc


def test_package_imports_and_reports_a_version() -> None:
    assert isinstance(tex2gdoc.__version__, str)
    assert tex2gdoc.__version__
