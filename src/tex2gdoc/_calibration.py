"""Calibration cases: proof that each check fires on input built to break it.

A check that has never been seen to fail is not evidence. Twelve of this
converter's checks had never failed on anything, which is exactly the condition
under which a check that inspects nothing looks identical to one that works.

Every case here is built in memory. `verify()` takes three plain arguments, a
path to a zip, a `SourceFacts`, and citeproc's stderr as a string, so none of
this needs pandoc or a TeX engine and all of it runs in CI.

Two properties are enforced, and the second is the one that is easy to forget:

  * every check must reach its worst status on input built to break it, and
  * the known-good case must emit exactly the expected SET of check names.

The second exists because several checks are conditional on a `SourceFacts`
field being non-empty. If a source-parsing bug empties `listing_probes`, the
listing checks are not reported as failing, they are not reported at all, and a
run that only looks for FAIL calls that clean.
"""

from __future__ import annotations

import ast
import inspect
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from .tex2gdoc import MARGIN, TEXT_WIDTH, SourceFacts
from .verification import verify


def declared_check_names() -> set[str]:
    """Every check name `verify()` can emit, read from its own source.

    Static rather than dynamic on purpose: a check guarded by
    `if facts.listing_probes:` has to be discovered without being executed, or
    the coverage gate would only cover whichever branches a sample input
    happened to take.

    Relies on one convention, worth keeping: a check name is a literal string,
    never an f-string.

    The whole module is parsed, not just `verify`. A check built in a helper
    that `verify` calls is still a check, and reading only `verify` would let
    one skip the coverage gate entirely: the exact condition this file exists
    to prevent, reintroduced by where someone chose to put a function.
    """
    tree = ast.parse(inspect.getsource(sys.modules[verify.__module__]))
    return {
        node.args[0].value
        for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id == "Check"
        and node.args
        and isinstance(node.args[0], ast.Constant)
        and isinstance(node.args[0].value, str)
    }


# --------------------------------------------------------------------------
# A known-good pair of XML parts, built from the module's own constants so the
# fixture cannot drift away from the code it is calibrating.
# --------------------------------------------------------------------------

CAPTION_PROBE = "a distinctive caption phrase"
TABLE_PROBE = "Provenance"
LISTING_PROBE = "structure Loop where"
CITATION_KEY = "present2020"

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_COL = TEXT_WIDTH // 2
_COLS = [_COL, TEXT_WIDTH - _COL]


def _para(text: str, style: str | None = None) -> str:
    pstyle = f'<w:pStyle w:val="{style}" />' if style else ""
    return f'<w:p><w:pPr>{pstyle}</w:pPr><w:r><w:t xml:space="preserve">{text} </w:t></w:r></w:p>'


def _cell(text: str, width: int) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{width}" /></w:tcPr>{_para(text)}</w:tc>'


def _table() -> str:
    grid = "".join(f'<w:gridCol w:w="{w}" />' for w in _COLS)
    header = "".join(_cell(t, w) for t, w in zip(["Claim", TABLE_PROBE], _COLS, strict=True))
    row = "".join(_cell(t, w) for t, w in zip(["Alpha", "Bravo"], _COLS, strict=True))
    return (
        f'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="{TEXT_WIDTH}" />'
        '<w:tblLayout w:type="fixed" /></w:tblPr>'
        f"<w:tblGrid>{grid}</w:tblGrid>"
        f"<w:tr>{header}</w:tr><w:tr>{row}</w:tr></w:tbl>"
    )


def _code_paragraph(text: str, coloured: bool = True) -> str:
    # The exact adjacency the "code highlighted" detector looks for. It is
    # written out in full here rather than assembled, because the detector
    # depends on the literal spacing and a fixture that quietly differs would
    # calibrate nothing.
    colour = '<w:color w:val="0033B3" />' if coloured else ""
    return (
        '<w:p><w:pPr><w:pStyle w:val="SourceCode" /></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="VerbatimChar" />'
        f"{colour}</w:rPr>"
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


_SECT_PR = (
    f'<w:sectPr><w:pgSz w:w="12240" w:h="15840" />'
    f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}"'
    f' w:left="{MARGIN}" /></w:sectPr>'
)

_BIB_ENTRY = (
    "Example, Ann. A Present Reference. Journal of Testing, 2020. "
    "Pages 1 to 12, with enough text to clear the length floor."
)


def good_document_xml(
    *,
    body: str | None = None,
    body_style: str | None = None,
    omath: int = 2,
    heading_style: str = "Heading1",
    code_coloured: bool = True,
    code_text: str = LISTING_PROBE,
    caption: str = CAPTION_PROBE,
    table: str | None = None,
    references: str = "References",
    references_style: str = "Heading1",
    bib_entry: str = _BIB_ENTRY,
) -> str:
    """Build a document.xml. Every keyword is a lever a calibration case pulls."""
    math = "".join("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>" for _ in range(omath))
    prose = (
        body
        if body is not None
        else (
            "This paragraph carries enough ordinary prose that the body text volume "
            "check has something to measure against the source it came from. "
        )
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        "<w:body>"
        + _para("Section One", heading_style)
        + (
            f"<w:p><w:pPr>{f"<w:pStyle w:val='{body_style}' />" if body_style else ''}</w:pPr>"
            f'<w:r><w:t xml:space="preserve">{prose}</w:t></w:r>{math}</w:p>'
        )
        + _para(caption, "ImageCaption")
        + (table if table is not None else _table())
        + _code_paragraph(code_text, coloured=code_coloured)
        + _para(references, references_style)
        # Normal, not pandoc Bibliography: the converter remaps that one, so a
        # fixture using it would no longer describe what ships.
        + _para(bib_entry)
        + _SECT_PR
        + "</w:body></w:document>"
    )


def good_styles_xml(
    *,
    captioned_figure_centered: bool = True,
    image_caption_centered: bool = True,
    trailing_centered_style: bool = False,
) -> str:
    """Build a styles.xml.

    `trailing_centered_style` adds an unrelated centred style *after* the figure
    styles. It exists to catch a specific defect: an unanchored DOTALL search for
    a centred justification anywhere after a style id will match a later style's
    centring and report the earlier, uncentred one as fine.
    """

    def style(sid: str, centered: bool) -> str:
        jc = '<w:jc w:val="center" />' if centered else '<w:jc w:val="left" />'
        return (
            f'<w:style w:type="paragraph" w:styleId="{sid}">'
            f'<w:name w:val="{sid}" /><w:pPr>{jc}</w:pPr></w:style>'
        )

    parts = [
        style("CaptionedFigure", captioned_figure_centered),
        style("ImageCaption", image_caption_centered),
        style("BodyText", False),
    ]
    if trailing_centered_style:
        parts.append(style("SomethingElse", True))
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "".join(parts)
        + "</w:styles>"
    )


def good_facts(document_xml: str) -> SourceFacts:
    """Facts that match the good document.

    `body_chars` is derived from the document so the volume check starts at a
    ratio of 1.0 and a case that truncates the body moves it decisively.
    """
    import re

    text = re.sub(r"<[^>]+>", "", document_xml)
    return SourceFacts(
        tikz_figures=1,
        graphics_figures=0,
        top_level_tabulars=1,
        table_probes=[TABLE_PROBE],
        caption_probes=[CAPTION_PROBE],
        listing_probes=[LISTING_PROBE],
        citation_keys={CITATION_KEY},
        inline_math=2,
        sections=1,
        body_chars=len(text),
    )


# A complete, clean package for the known-good case. Without these the three
# package-level checks pass against an absence: no .rels to walk, no settings
# to find an embed flag in, no theme and no font table. That is the condition
# this module exists to prevent, reproduced inside the module itself.
_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"

GOOD_RELS = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    f'<Relationships xmlns="{_PKG_RELS}">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
    ' Target="https://example.invalid/paper" TargetMode="External" />'
    '<Relationship Id="rId2"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
    ' Target="media/image0.png" />'
    "</Relationships>"
)

GOOD_SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    f'<w:settings xmlns:w="{_W}"><w:zoom w:percent="100" /></w:settings>'
)

GOOD_THEME = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    f'<a:theme xmlns:a="{_DRAWINGML}"><a:themeElements>'
    '<a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="" /><a:ea typeface="" /><a:cs typeface="" /></a:majorFont>'
    '<a:minorFont><a:latin typeface="" /><a:ea typeface="" /><a:cs typeface="" /></a:minorFont>'
    "</a:fontScheme></a:themeElements></a:theme>"
)

GOOD_FONT_TABLE = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    f'<w:fonts xmlns:w="{_W}"><w:font w:name="Courier New" />'
    '<w:font w:name="Cambria Math" /></w:fonts>'
)

GOOD_PACKAGE = {
    "word/_rels/document.xml.rels": GOOD_RELS,
    "word/settings.xml": GOOD_SETTINGS,
    "word/theme/theme1.xml": GOOD_THEME,
    "word/fontTable.xml": GOOD_FONT_TABLE,
}


def write_docx(
    directory: Path,
    document_xml: str,
    styles_xml: str,
    media: int = 1,
    name: str = "case.docx",
    extra_parts: dict[str, str] | None = None,
    package: bool = True,
) -> Path:
    out = directory / name
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", styles_xml)
        for i in range(media):
            z.writestr(f"word/media/image{i}.png", b"\x89PNG\r\n\x1a\n")
        # A complete clean package by default, so the package-level checks are
        # exercised against something correct rather than against nothing.
        # `extra_parts` overrides an entry to make one of them fail.
        parts = {**(GOOD_PACKAGE if package else {}), **(extra_parts or {})}
        for part, body in parts.items():
            z.writestr(part, body)
    return out


# --------------------------------------------------------------------------
# The registry
# --------------------------------------------------------------------------

Build = Callable[[Path], tuple[Path, SourceFacts, str]]


@dataclass(frozen=True)
class Calibration:
    """One check, and input built to make it report `worst`."""

    check: str
    worst: str
    build: Build
    why: str


def _case(
    *,
    document: str | None = None,
    styles: str | None = None,
    media: int = 1,
    stderr: str = "",
    facts_from: str | None = None,
    mutate_facts: Callable[[SourceFacts], None] | None = None,
    extra_parts: dict[str, str] | None = None,
    package: bool = True,
) -> Build:
    """Assemble a build function from the levers above."""

    def build(tmp: Path) -> tuple[Path, SourceFacts, str]:
        doc = document if document is not None else good_document_xml()
        sty = styles if styles is not None else good_styles_xml()
        facts = good_facts(facts_from if facts_from is not None else good_document_xml())
        if mutate_facts:
            mutate_facts(facts)
        return (
            write_docx(tmp, doc, sty, media=media, extra_parts=extra_parts, package=package),
            facts,
            stderr,
        )

    return build


# A relationship part carrying one image the reader's machine would go and
# fetch on open. The Target is unreachable on purpose: nothing in the test
# suite may touch the network, and the check reads the declaration, not the URL.
LINKED_IMAGE_RELS = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId99"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
    ' Target="https://example.invalid/logo.png" TargetMode="External" />'
    "</Relationships>"
)

# A theme naming a face in both slots, the state pandoc's reference doc arrives
# in. The east-Asian and complex-script slots are left empty, which is both what
# a real theme looks like and a second thing the check must not mistake for a
# pinned font.
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
THEME_NAMING_A_FONT = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    f'<a:theme xmlns:a="{_A}"><a:themeElements><a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204" />'
    '<a:ea typeface="" /><a:cs typeface="" /></a:majorFont>'
    '<a:minorFont><a:latin typeface="" /><a:font script="Jpan" typeface="游ゴシック" />'
    '<a:ea typeface="" /><a:cs typeface="" /></a:minorFont>'
    "</a:fontScheme></a:themeElements></a:theme>"
)


# Checks that return early and therefore never appear beside the others.
EARLY_RETURN_CHECKS = {"output exists", "document parses"}


def _missing_output(tmp: Path) -> tuple[Path, SourceFacts, str]:
    return tmp / "never-written.docx", good_facts(good_document_xml()), ""


def _unparseable_output(tmp: Path) -> tuple[Path, SourceFacts, str]:
    """A .docx whose document.xml is not well-formed XML."""
    docx = write_docx(tmp, "<w:document><w:body><w:p></w:document>", good_styles_xml())
    return docx, good_facts(good_document_xml()), ""


def _source_was_much_longer(facts: SourceFacts) -> None:
    """A source far larger than what arrived, which is what truncation looks like."""
    facts.body_chars = 100_000


CALIBRATIONS: dict[str, Calibration] = {
    "output exists": Calibration(
        "output exists",
        "FAIL",
        _missing_output,
        "no .docx at the path at all; the guard that stops every later check "
        "from passing vacuously on an absent file",
    ),
    "document parses": Calibration(
        "document parses",
        "FAIL",
        _unparseable_output,
        "a .docx whose document.xml is not well-formed; the tree rewrite cannot "
        "produce one, and this is what says so if it ever does",
    ),
    "figures embedded": Calibration(
        "figures embedded",
        "FAIL",
        _case(media=0),
        "the source declares a figure and the archive carries no image",
    ),
    "tables present": Calibration(
        "tables present",
        "FAIL",
        _case(document=good_document_xml(table="")),
        "the source declares a tabular and the output has no table element",
    ),
    "table layout": Calibration(
        "table layout",
        "FAIL",
        _case(document=good_document_xml().replace('w:type="dxa"', 'w:type="pct"')),
        "percentage widths, which Google Docs renders unreliably",
    ),
    "column widths sum": Calibration(
        "column widths sum",
        "FAIL",
        _case(
            document=good_document_xml().replace(
                f'<w:gridCol w:w="{_COLS[0]}" />', '<w:gridCol w:w="10" />'
            )
        ),
        "a grid whose columns no longer add up to the table width",
    ),
    "figures centered": Calibration(
        "figures centered",
        "FAIL",
        _case(
            styles=good_styles_xml(captioned_figure_centered=False, trailing_centered_style=True)
        ),
        "the figure style is not centred, while a later unrelated style is; a "
        "DOTALL search that is not scoped to one style block reports this as fine",
    ),
    "page geometry": Calibration(
        "page geometry",
        "FAIL",
        _case(document=good_document_xml().replace(f'w:left="{MARGIN}"', 'w:left="99"')),
        "the page margins were never applied",
    ),
    "code highlighted": Calibration(
        "code highlighted",
        "FAIL",
        _case(document=good_document_xml(code_coloured=False)),
        "a code listing that reached the output with no colouring",
    ),
    "table contents survived": Calibration(
        "table contents survived",
        "FAIL",
        _case(document=good_document_xml(table=_table().replace(TABLE_PROBE, "gone"))),
        "an empty-shell table: the element count still passes, the content is lost",
    ),
    "captions survived": Calibration(
        "captions survived",
        "FAIL",
        _case(document=good_document_xml(caption="something else entirely")),
        "the caption of a starred float, dropped by pandoc without warning",
    ),
    "code listings readable": Calibration(
        "code listings readable",
        "FAIL",
        _case(document=good_document_xml(code_text="structureLoopwhere")),
        "the alltt collapse: every run of spaces gone from a listing",
    ),
    "cross-references resolved": Calibration(
        "cross-references resolved",
        "FAIL",
        _case(document=good_document_xml(body="see [tab:map] for the mapping ")),
        "a raw label key left in the text where a number belongs",
    ),
    "citations resolved": Calibration(
        "citations resolved",
        "FAIL",
        _case(stderr="[WARNING] Citeproc: citation missing1999 not found"),
        "a cited key with no bibliography entry",
    ),
    "bibliography rendered": Calibration(
        "bibliography rendered",
        "FAIL",
        _case(document=good_document_xml(references="Nothing Here", bib_entry="")),
        "no reference section, or one with no entries under it",
    ),
    "math converted": Calibration(
        "math converted",
        "FAIL",
        _case(document=good_document_xml(omath=0)),
        "math in the source and not one equation in the output",
    ),
    "sections present": Calibration(
        "sections present",
        "FAIL",
        # Both headings have to go. Leaving the References heading as Heading1
        # kept the count at 1, which met the threshold and reported PASS: the
        # first version of this case proved nothing.
        _case(document=good_document_xml(heading_style="BodyText", references_style="BodyText")),
        "the heading styles were never applied, so the document has no outline",
    ),
    "body text volume": Calibration(
        "body text volume",
        "FAIL",
        # Truncating the body alone was not enough: the table, listing and
        # bibliography still carry text, so the ratio stayed above the floor.
        # Stating that the source was far longer is what truncation actually is.
        _case(document=good_document_xml(body="x"), mutate_facts=_source_was_much_longer),
        "most of the prose missing, against a source that had it",
    ),
    "no external dependencies": Calibration(
        "no external dependencies",
        "FAIL",
        # An externally linked image, which is the family that matters: opening
        # the file reaches out to a host the reader never chose. A hyperlink
        # would not do here, and a case built from one would have calibrated
        # nothing, since hyperlinks are the allowed kind.
        _case(extra_parts={"word/_rels/document.xml.rels": LINKED_IMAGE_RELS}),
        "an image the reader's machine would fetch from the network on open",
    ),
    "body styles are native": Calibration(
        "body styles are native",
        "FAIL",
        # A body paragraph left on pandoc's BodyText, which Docs flattens into
        # direct formatting that its style menu cannot reach.
        _case(document=good_document_xml(body_style="BodyText")),
        "running prose on a style Google Docs turns into direct formatting",
    ),
    "fonts left to the template": Calibration(
        "fonts left to the template",
        "FAIL",
        # A theme that names a face, which is what pandoc's reference doc ships
        # and what overrides the template the file is imported into. A monospace
        # pin would not do: that one is allowed, so a case built from it would
        # calibrate nothing.
        _case(extra_parts={"word/theme/theme1.xml": THEME_NAMING_A_FONT}),
        "the theme names a typeface, overriding whatever template hosts the file",
    ),
}


def run_calibration(tmp: Path, verbose: bool = False) -> tuple[bool, list[str]]:
    """Run the coverage gate and every calibration case. Returns (ok, lines)."""
    lines: list[str] = []
    ok = True

    declared = declared_check_names()
    registered = set(CALIBRATIONS)
    for missing in sorted(declared - registered):
        ok = False
        lines.append(f"  [FAIL] no calibration case for check {missing!r}")
    for extra in sorted(registered - declared):
        ok = False
        lines.append(f"  [FAIL] calibration for {extra!r}, which verify() no longer emits")
    if declared == registered:
        lines.append(f"  [PASS] coverage gate: all {len(declared)} checks have a case")

    # Known-good input: no failures, and exactly the expected set of names.
    good_doc = good_document_xml()
    good_dir = tmp / "good"
    good_dir.mkdir(parents=True, exist_ok=True)
    good = write_docx(good_dir, good_doc, good_styles_xml(), name="good.docx")
    checks = verify(good, good_facts(good_doc), "")
    emitted = {c.name for c in checks}
    failures = [c for c in checks if c.status != "PASS"]
    expected = declared - EARLY_RETURN_CHECKS
    if failures:
        ok = False
        lines.append(
            "  [FAIL] known-good input: " + ", ".join(f"{c.name}={c.status}" for c in failures)
        )
    elif emitted != expected:
        ok = False
        lines.append(
            f"  [FAIL] known-good input emitted {sorted(emitted)}, expected {sorted(expected)}"
        )
    else:
        lines.append(f"  [PASS] known-good input: all {len(emitted)} applicable checks PASS")

    for name in sorted(CALIBRATIONS):
        cal = CALIBRATIONS[name]
        case_dir = tmp / name.replace(" ", "-")
        case_dir.mkdir(parents=True, exist_ok=True)
        docx, facts, stderr = cal.build(case_dir)
        got = next((c for c in verify(docx, facts, stderr) if c.name == name), None)
        if got is None:
            ok = False
            lines.append(f"  [FAIL] {name}: not emitted at all for its own broken input")
        elif got.status != cal.worst:
            ok = False
            lines.append(f"  [FAIL] {name}: {got.status}, wanted {cal.worst} ({cal.why})")
        elif verbose:
            lines.append(f"  [PASS] {name}: {cal.worst} on {cal.why}")

    return ok, lines
