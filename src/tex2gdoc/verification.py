"""Check the .docx that ships against what the source claimed.

Separate from the converter on purpose. This is the layer that decides whether a
run can be trusted, and it reads the output file rather than any intermediate.

Every check here is calibrated: `tex2gdoc._calibration` builds input designed to
break each one and asserts that it reports FAIL, and a check added without a
calibration case fails the coverage gate. Silence is not evidence.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Final

from .ooxml import (
    Attr,
    SafeElement,
    Tag,
    UnsafePart,
    Value,
    attribute,
    child,
    descendants,
    parse_part,
    value_of,
)
from .tex2gdoc import CODE_FONT, MARGIN, MATH_FONT, TEXT_WIDTH, SourceFacts

# The one family a document must name, because OOXML has no theme slot for it
# and a listing set in the body face is no longer a listing.
# The only two names allowed to survive. CODE_FONT is the one the converter
# writes for listings, because no theme slot supplies monospace; MATH_FONT is
# the sole surviving hint about how the equations are meant to be laid out.
# A wider allowlist would pass a document pinning Menlo, which is unreadable as
# intended anywhere but macOS: the exact portability failure this check names.
ALLOWED_FONT_NAMES: Final = frozenset({CODE_FONT, MATH_FONT})

FONT_SLOTS: Final = (Attr.ASCII, Attr.HIGH_ANSI, Attr.COMPLEX_SCRIPT, Attr.EAST_ASIAN)
FONT_EMBED_TAGS: Final = (
    Tag.EMBED_REGULAR,
    Tag.EMBED_BOLD,
    Tag.EMBED_ITALIC,
    Tag.EMBED_BOLD_ITALIC,
)
# Field codes that make a renderer go and fetch something when the file opens.
FETCHING_FIELDS: Final = ("INCLUDEPICTURE", "INCLUDETEXT")
# Paragraph styles Docs has no equivalent for, used here for running prose.
# The styles STYLE_REMAP moves. Kept in step with it: a name here that the
# converter no longer remaps would fail every clean document.
UNMAPPABLE_BODY_STYLES: Final = ("BodyText", "FirstParagraph", "Compact", "Bibliography")


class Status(StrEnum):
    """A check's verdict. `run_calibration` compares these, so they are a set."""

    PASS = "PASS"  # noqa: S105 - a check verdict, not a credential
    WARN = "WARN"
    FAIL = "FAIL"


@dataclass
class Check:
    name: str
    status: Status
    detail: str


@dataclass(frozen=True)
class Package:
    """Every part `verify` reads, opened once and parsed inside one guarded read.

    The checks used to reopen the zip for themselves. That parsed the largest
    part twice and, worse, put those parses outside the try/except that turns an
    unreadable file into `document parses: FAIL`, so a co-author's .docx with a
    malformed settings.xml took the CLI down with a traceback instead. `verify`
    now reports every way the file can be unreadable, including one that is not
    a zip at all, rather than raising for some of them.
    """

    document: SafeElement
    text: str
    media: list[str]
    names: list[str]
    styles: SafeElement | None
    settings: SafeElement | None
    font_table: SafeElement | None
    theme: SafeElement | None
    relationship_parts: list[SafeElement]


def read_docx(path: Path) -> Package:
    """Open the .docx once and parse every part the checks need.

    Parsed rather than matched. Every detector below reads a tree, so a change
    in how pandoc spaces its tags cannot move a count, and a malformed part
    raises here instead of quietly matching nothing.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        def part(name: str) -> SafeElement | None:
            if name not in names:
                return None
            # Bytes, not str: decoding with errors="replace" first would let a
            # part be silently repaired before inspection, so the checks would
            # read something Word never sees. The parser handles the encoding
            # declaration itself, and rejects what it cannot decode.
            return parse_part(z.read(name), name)

        document = part("word/document.xml")
        if document is None:
            raise KeyError("word/document.xml")
        relationships = [e for n in names if n.endswith(".rels") if (e := part(n)) is not None]
        return Package(
            document=document,
            text="".join(t for t in document.itertext() if isinstance(t, str)),
            media=[n for n in names if n.startswith("word/media/")],
            names=names,
            styles=part("word/styles.xml"),
            settings=part("word/settings.xml"),
            font_table=part("word/fontTable.xml"),
            theme=part("word/theme/theme1.xml"),
            relationship_parts=relationships,
        )


def external_dependencies(package: Package) -> Check:
    """Fail if opening the .docx would make the reader's machine fetch anything.

    A review copy is passed around and opened by people who did not build it, so
    it has to be self-contained: every byte it renders is a byte inside the zip.
    Two families are caught. Resources a renderer fetches at open time (linked
    images, subdocuments, an attached template, any non-hyperlink external
    relationship), and fonts, whether pulled in or embedded, because an embedded
    font ships a third party's binary under someone else's licence.

    Hyperlinks are counted and allowed: a bibliography is mostly DOIs, and a
    hyperlink is followed on click, never on open. That distinction is the whole
    point of the check, so it is stated in the detail line rather than left for
    a reader to infer from a bare PASS.

    Scope, because a green line here is easy to over-read: this asks whether
    *this converter's own output* reaches the network when opened. It is not a
    safety verdict on a document from somewhere else. It does not look for
    embedded OLE objects, ActiveX controls or macros, and a PASS on a file this
    tool did not write says nothing about any of them.
    """
    offenders: list[str] = []
    hyperlinks = 0

    for rels in package.relationship_parts:
        for rel in descendants(rels, Tag.RELATIONSHIP):
            if attribute(rel, Attr.TARGET_MODE) != Value.EXTERNAL:
                continue
            kind = (attribute(rel, Attr.RELATIONSHIP_TYPE) or "").rsplit("/", 1)[-1]
            if kind == Value.HYPERLINK:
                hyperlinks += 1
            else:
                offenders.append(
                    f"external {kind or 'unknown'} -> {attribute(rel, Attr.TARGET) or '?'}"
                )

    offenders += [
        f"embedded font part {n}"
        for n in package.names
        if n.lower().endswith((".odttf", ".ttf", ".otf"))
    ]

    document = package.document
    offenders += ["subdocument reference" for _ in descendants(document, Tag.SUBDOCUMENT)]
    # A linked image carries the relationship id in r:link (DrawingML) or
    # r:href (VML); r:embed, the id of a part inside the zip, is the
    # self-contained spelling and is what should be there instead.
    for element in (SafeElement(e) for e in document.iter()):
        if attribute(element, Attr.LINK) or attribute(element, Attr.HREF):
            offenders.append("linked image (fetched on open, not embedded)")
    for instruction in [
        *(e.text or "" for e in descendants(document, Tag.INSTRUCTION_TEXT)),
        *(attribute(e, Attr.INSTRUCTION) or "" for e in descendants(document, Tag.SIMPLE_FIELD)),
    ]:
        # Word accepts a field name in any case, so a lowercase INCLUDEPICTURE
        # fetches on open exactly the same and must not read as clean.
        upper = instruction.upper()
        if any(field in upper for field in FETCHING_FIELDS):
            offenders.append("INCLUDEPICTURE/INCLUDETEXT field")

    if package.settings is not None:
        for flag in (Tag.EMBED_SYSTEM_FONTS, Tag.EMBED_TRUETYPE_FONTS):
            if child(package.settings, flag) is not None:
                offenders.append(f"font-embedding directive {flag}")
        if child(package.settings, Tag.ATTACHED_TEMPLATE) is not None:
            offenders.append("attached template")

    if package.font_table is not None and any(
        descendants(package.font_table, tag) for tag in FONT_EMBED_TAGS
    ):
        offenders.append("font embedding in fontTable.xml")

    allowed = f"{hyperlinks} hyperlink(s), followed on click not on open"
    return Check(
        "no external dependencies",
        Status.PASS if not offenders else Status.FAIL,
        f"nothing is fetched or embedded on open; {allowed}"
        if not offenders
        else f"{len(offenders)} external dependency(ies): {'; '.join(sorted(set(offenders))[:4])}",
    )


def fonts_left_to_the_template(package: Package) -> Check:
    """Fail if the document names a typeface it has no need to name.

    A review copy gets pasted into someone else's template, and every font this
    file names is one the template asked for and did not get. The styles are
    expected to resolve through the theme, the theme is expected to name
    nothing, and the font table is expected to describe only what is left.

    One exception is allowed and it is not the maths. Equations carry no font
    here at all, so a renderer reaches for its own maths face without being
    told. Monospace is the exception: a code listing set in a proportional font
    stops being a listing, and no theme slot means monospace.
    """
    named: set[str] = set()

    if package.theme is not None:
        for slot in (Tag.MAJOR_FONT, Tag.MINOR_FONT):
            for scheme in descendants(package.theme, slot):
                where = slot.removeprefix("a:")
                # The whole scheme, not the latin slot. A theme names a face per
                # script as well, 94 of them in pandoc's, and a detector that
                # reads three attributes reports a file with 94 typefaces in it
                # as naming none.
                for element in (SafeElement(e) for e in scheme.iter()):
                    if typeface := attribute(element, Attr.TYPEFACE):
                        script = attribute(element, Attr.SCRIPT) or "latin"
                        named.add(f"theme {where} [{script}]: {typeface}")
                    # A PANOSE left behind still steers substitution by metrics,
                    # so a cleared name alone is not the guarantee this claims.
                    if panose := attribute(element, Attr.PANOSE):
                        named.add(f"theme {where} panose: {panose}")

    # The font table is the other place a name survives. Leaving it out is how
    # this check reported a file still advertising Aptos as naming nothing.
    if package.font_table is not None:
        for font in descendants(package.font_table, Tag.FONT):
            name = attribute(font, Attr.NAME)
            if name and name not in ALLOWED_FONT_NAMES:
                named.add(f"font table: {name}")

    for tree in (package.document, package.styles):
        if tree is None:
            continue
        for fonts in descendants(tree, Tag.FONTS):
            # Theme-relative attributes (w:asciiTheme and friends) are the
            # wanted spelling and are deliberately not collected here.
            literal = {attribute(fonts, slot) for slot in FONT_SLOTS}
            named |= {f for f in literal if f and f not in ALLOWED_FONT_NAMES}

    return Check(
        "fonts left to the template",
        Status.PASS if not named else Status.FAIL,
        "no typeface named beyond the code and maths faces; the theme is empty"
        if not named
        else f"{len(named)} typeface(s) pinned: {'; '.join(sorted(named)[:4])}",
    )


def body_styles_are_native(document: SafeElement) -> Check:
    """Fail if running prose still carries a style Google Docs cannot map.

    Docs turns an unmapped paragraph style into direct formatting, which its
    style menu cannot reach, so the reader can restyle headings and nothing
    else. The converter moves those paragraphs onto `Normal`; this is what
    notices when a pandoc change renames them and the remap silently stops
    matching.

    Scope is exactly what `STYLE_REMAP` moves, which is the set that provably
    renders the same on either style. `Abstract`, `BlockText` and the captions
    keep their own styles because remapping them would move something on the
    page, so this check says nothing about them.
    """
    stranded = sorted(
        {
            value
            for pstyle in descendants(document, Tag.PARAGRAPH_STYLE)
            if (value := value_of(pstyle)) in UNMAPPABLE_BODY_STYLES
        }
    )
    return Check(
        "body styles are native",
        Status.PASS if not stranded else Status.FAIL,
        "running prose is on Normal, which Docs maps to its own body style"
        if not stranded
        else f"{len(stranded)} body style(s) Docs cannot map: {', '.join(stranded)}",
    )


def verify(docx: Path, facts: SourceFacts, citeproc_stderr: str) -> list[Check]:
    """Check the output against what the source claimed. Silence is not evidence."""
    checks: list[Check] = []

    if not docx.exists() or docx.stat().st_size == 0:
        return [Check("output exists", Status.FAIL, "no .docx was written")]

    # Reading the output is now a parse, so a malformed part is caught here
    # rather than silently matching nothing further down. That is most of what
    # a schema validator was wanted for: this pass cannot write a document that
    # will not parse, and it says so out loud if one ever arrives.
    try:
        package = read_docx(docx)
    except (UnsafePart, KeyError, zipfile.BadZipFile, OSError) as exc:
        # BadZipFile and OSError matter as much as a malformed part: `verify`
        # is documented as reading a file that came back from a co-author, and
        # one that is not a .docx at all used to reach them as a traceback.
        return [Check("document parses", Status.FAIL, f"the .docx is not readable: {exc}")]

    document, text = package.document, package.text
    media, styles = package.media, package.styles
    flat = " ".join(text.split())

    expected_images = facts.tikz_figures + facts.graphics_figures
    checks.append(
        Check(
            "figures embedded",
            Status.PASS if len(media) >= expected_images else Status.FAIL,
            f"{len(media)} image(s) in the .docx, {expected_images} figure(s) in the source",
        )
    )

    tables = len(list(descendants(document, Tag.TABLE)))
    checks.append(
        Check(
            "tables present",
            Status.PASS if tables >= facts.top_level_tabulars else Status.FAIL,
            f"{tables} table(s) in the .docx, {facts.top_level_tabulars} tabular(s) in the source",
        )
    )

    if tables:
        full = sum(
            1
            for e in descendants(document, Tag.TABLE_WIDTH)
            if attribute(e, Attr.TYPE) == Value.DXA and attribute(e, Attr.WIDTH) == str(TEXT_WIDTH)
        )
        cells = len(list(descendants(document, Tag.TABLE_CELL)))
        sized = sum(
            1 for e in descendants(document, Tag.CELL_WIDTH) if attribute(e, Attr.TYPE) == Value.DXA
        )
        pct = sum(
            1
            for e in (SafeElement(x) for x in document.iter())
            if attribute(e, Attr.TYPE) == Value.PERCENT
        )
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
                Status.PASS if not problems else Status.FAIL,
                f"{tables} table(s) at {TEXT_WIDTH} dxa, {cells} cell(s) sized in dxa"
                if not problems
                else "; ".join(problems),
            )
        )

        bad = [
            g
            for g in descendants(document, Tag.TABLE_GRID)
            if sum(int(attribute(c, Attr.WIDTH) or 0) for c in descendants(g, Tag.GRID_COLUMN))
            != TEXT_WIDTH
        ]
        checks.append(
            Check(
                "column widths sum",
                Status.PASS if not bad else Status.FAIL,
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
            for style in descendants(styles, Tag.STYLE):
                if attribute(style, Attr.STYLE_ID) != sid:
                    continue
                ppr = child(style, Tag.PARAGRAPH_PROPERTIES)
                return ppr is not None and value_of(child(ppr, Tag.JUSTIFICATION)) == "center"
            return False

        centered = all(style_is_centered(sid) for sid in ("CaptionedFigure", "ImageCaption"))
        checks.append(
            Check(
                "figures centered",
                Status.PASS if centered else Status.FAIL,
                "figure and caption styles are centered"
                if centered
                else "figure or caption style is not centered — did the reference document build?",
            )
        )

    margins = descendants(document, Tag.PAGE_MARGIN)
    margins_applied = bool(margins) and attribute(margins[0], Attr.LEFT) == str(MARGIN)
    checks.append(
        Check(
            "page geometry",
            Status.PASS if margins_applied else Status.FAIL,
            f"margins set to {MARGIN} dxa"
            if margins_applied
            else "page margins were not applied — did the reference document build?",
        )
    )

    if facts.listing_probes:
        colored = sum(
            1
            for run in descendants(document, Tag.RUN)
            if (rpr := child(run, Tag.RUN_PROPERTIES)) is not None
            and value_of(child(rpr, Tag.RUN_STYLE)) == "VerbatimChar"
            and child(rpr, Tag.COLOR) is not None
        )
        # FAIL, not WARN. report() exits nonzero only on FAIL, so a WARN here
        # meant a refactor could silently drop every coloured run and still
        # leave the command green.
        checks.append(
            Check(
                "code highlighted",
                Status.PASS if colored else Status.FAIL,
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
                Status.PASS if not lost else Status.FAIL,
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
                Status.PASS if not lost else Status.FAIL,
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
                Status.PASS if not lost else Status.FAIL,
                f"{len(facts.listing_probes)} listing(s) probed, spacing intact"
                if not lost
                else f"{len(lost)} listing(s) lost their spacing, first: {lost[0]!r}",
            )
        )

    dangling = sorted(set(re.findall(r"\[(?:tab|fig|thm|sec|app|eq):[A-Za-z0-9_\-]+\]", text)))
    checks.append(
        Check(
            "cross-references resolved",
            Status.PASS if not dangling else Status.FAIL,
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
                Status.PASS if not missing else Status.FAIL,
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
                Status.PASS if pos >= 0 and tail >= floor else Status.FAIL,
                f"reference section found, {tail} chars of entries (floor {floor})"
                if pos >= 0
                else "no reference section — was a bibliography passed?",
            )
        )

    math = len(list(descendants(document, Tag.MATH)))
    if facts.inline_math:
        ratio = math / facts.inline_math
        status = Status.FAIL if math == 0 else (Status.WARN if ratio < 0.5 else Status.PASS)
        checks.append(
            Check(
                "math converted",
                status,
                f"{math} Word equation(s), {facts.inline_math} inline-math span(s) in the source "
                f"(ratio {ratio:.2f}; exact parity is not expected)",
            )
        )

    headings = sum(
        1 for e in descendants(document, Tag.PARAGRAPH_STYLE) if value_of(e) == "Heading1"
    )
    if facts.sections:
        # Also FAIL rather than WARN: fewer top-level headings than the source
        # has sections means the outline a reviewer navigates by is incomplete,
        # and that is not a warning.
        checks.append(
            Check(
                "sections present",
                Status.PASS if headings >= facts.sections else Status.FAIL,
                f"{headings} top-level heading(s), {facts.sections} \\section(s) in the source",
            )
        )

    ratio = len(text) / facts.body_chars if facts.body_chars else 0
    checks.append(
        Check(
            "body text volume",
            Status.PASS if ratio >= 0.5 else Status.FAIL,
            f"{len(text)} chars extracted, {facts.body_chars} in the source (ratio {ratio:.2f})",
        )
    )

    checks.append(body_styles_are_native(document))
    checks.append(external_dependencies(package))
    checks.append(fonts_left_to_the_template(package))

    return checks
