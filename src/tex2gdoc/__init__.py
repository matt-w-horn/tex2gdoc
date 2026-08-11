"""Convert a LaTeX paper to a .docx for review in Word or Google Docs.

The converter lives in `tex2gdoc.tex2gdoc` and the checks in
`tex2gdoc.verification`. Both are re-exported here so callers can write
`from tex2gdoc import convert` rather than repeating the name.
`tex2gdoc.ooxml` sits under both: it owns the only XML parser, the element
builder, and the tag and attribute names, and it is not re-exported because
nothing outside the package should be building OOXML by hand.

Importing this package runs no external tool. `require_tools()` is called only
by `convert`'s CLI entry point and by `self_test`, so `import tex2gdoc` works
on a machine with no pandoc or TeX installed.
"""

from .tex2gdoc import (
    Figure,
    SourceFacts,
    collect_source_facts,
    convert,
    main,
    report,
    restyle_docx,
    self_test,
)
from .verification import Check, verify

__version__ = "0.1.0"

__all__ = [
    "Check",
    "Figure",
    "SourceFacts",
    "collect_source_facts",
    "convert",
    "main",
    "report",
    "restyle_docx",
    "self_test",
    "verify",
    "__version__",
]
