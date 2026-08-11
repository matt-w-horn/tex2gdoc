"""The only way this package is allowed to turn bytes into XML.

Three properties, and the type system carries all of them rather than a
convention that review has to keep enforcing.

Dangerous parsing is caught by annotation. `SafeElement` is a `NewType`, and
every function downstream takes one, so *passing* a tree from a bare
`etree.fromstring` is a mypy error at the call site rather than a runtime
surprise. What this does not do is make the wrong thing unrepresentable:
`SafeElement(...)` is an ordinary constructor, and this package itself calls it
to relabel `.iter()` results, which lxml types loosely. So the guard is a type
error on the easy mistake and a visible, greppable cast on the deliberate one.
The import ban below is what covers the rest.

Malformed parsing is loud. `recover=False` means a broken part raises instead of
yielding half a tree, which is the failure mode that lets a check inspect
nothing and report PASS.

Names are closed sets. `Tag`, `Attr` and `Value` are the vocabulary, so a
mistyped `w:spacng` is a mypy error rather than an element that silently
matches nothing. That is the same class of defect as the parser one: both fail
green. `Value` exists because `dxa` had been spelled as a literal in three
modules, and a keyword with three homes drifts.

The hardening itself is parser configuration rather than a scan of the bytes.
The DOCTYPE test that used to live here read the first 8 KB and admitted in its
own docstring that a long enough comment slid a declaration past the window.
`docinfo.doctype` is the parser's own answer, so there is no window.

None of this holds by agreement. ruff's banned-api rules reject importing any
XML parser anywhere but here: both `xml.etree` spellings, `xml.dom` and its two
submodules, `xml.sax`, `xml.parsers.expat`, `lxml.etree`, `lxml.objectify`,
`lxml.html` and `defusedxml`. So a second parser cannot appear without the lint
failing. It is still a list of names rather than a proof, so a parser that is
not on it is not caught: adding one to the project means adding it there too.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from enum import StrEnum
from typing import Final, NewType

from lxml import etree

# A prefix-to-URI mapping is what lxml's own API takes, so this one stays a
# dict. It is the only one in the module.
NS: Final[dict[str, str]] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    # The .rels parts, whose namespace is not the `r:` used for ids in a document.
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xml": "http://www.w3.org/XML/1998/namespace",
}


class Tag(StrEnum):
    """Every element this package reads or writes."""

    DOCUMENT = "w:document"
    BODY = "w:body"
    PARAGRAPH = "w:p"
    RUN = "w:r"
    TEXT = "w:t"
    RUN_PROPERTIES = "w:rPr"
    PARAGRAPH_PROPERTIES = "w:pPr"
    RUN_STYLE = "w:rStyle"
    PARAGRAPH_STYLE = "w:pStyle"
    STYLE = "w:style"
    COLOR = "w:color"
    BOLD = "w:b"
    ITALIC = "w:i"
    FONTS = "w:rFonts"
    FONT = "w:font"
    JUSTIFICATION = "w:jc"
    SPACING = "w:spacing"
    SHADING = "w:shd"
    KEEP_NEXT = "w:keepNext"
    KEEP_LINES = "w:keepLines"
    NUMBERING_PROPERTIES = "w:numPr"
    SECTION_PROPERTIES = "w:sectPr"
    PAGE_SIZE = "w:pgSz"
    PAGE_MARGIN = "w:pgMar"
    TABLE = "w:tbl"
    TABLE_ROW = "w:tr"
    TABLE_CELL = "w:tc"
    TABLE_PROPERTIES = "w:tblPr"
    TABLE_GRID = "w:tblGrid"
    GRID_COLUMN = "w:gridCol"
    TABLE_WIDTH = "w:tblW"
    TABLE_LAYOUT = "w:tblLayout"
    CELL_PROPERTIES = "w:tcPr"
    CELL_WIDTH = "w:tcW"
    GRID_SPAN = "w:gridSpan"
    INSTRUCTION_TEXT = "w:instrText"
    SIMPLE_FIELD = "w:fldSimple"
    SUBDOCUMENT = "w:subDoc"
    INDENT = "w:ind"
    DRAWING = "w:drawing"
    EMBED_SYSTEM_FONTS = "w:embedSystemFonts"
    EMBED_TRUETYPE_FONTS = "w:embedTrueTypeFonts"
    ATTACHED_TEMPLATE = "w:attachedTemplate"
    EMBED_REGULAR = "w:embedRegular"
    EMBED_BOLD = "w:embedBold"
    EMBED_ITALIC = "w:embedItalic"
    EMBED_BOLD_ITALIC = "w:embedBoldItalic"
    MATH = "m:oMath"
    MATH_PARA = "m:oMathPara"
    # OMML constructs that stand taller than a line of body text.
    NARY = "m:nary"
    FRACTION = "m:f"
    RADICAL = "m:rad"
    MATRIX = "m:m"
    SUB_SUPERSCRIPT = "m:sSubSup"
    LIMIT_LOWER = "m:limLow"
    LIMIT_UPPER = "m:limUpp"
    BOX = "m:box"
    MAJOR_FONT = "a:majorFont"
    MINOR_FONT = "a:minorFont"
    LATIN = "a:latin"
    RELATIONSHIP = "pr:Relationship"
    COMMENT = "w:comment"


class Attr(StrEnum):
    """Every attribute this package reads or writes."""

    VAL = "w:val"
    WIDTH = "w:w"
    HEIGHT = "w:h"
    TYPE = "w:type"
    TOP = "w:top"
    RIGHT = "w:right"
    BOTTOM = "w:bottom"
    LEFT = "w:left"
    HEADER = "w:header"
    FOOTER = "w:footer"
    GUTTER = "w:gutter"
    BEFORE = "w:before"
    AFTER = "w:after"
    LINE = "w:line"
    LINE_RULE = "w:lineRule"
    COLOR = "w:color"
    FILL = "w:fill"
    ASCII = "w:ascii"
    HIGH_ANSI = "w:hAnsi"
    COMPLEX_SCRIPT = "w:cs"
    EAST_ASIAN = "w:eastAsia"
    STYLE_ID = "w:styleId"
    NAME = "w:name"
    INSTRUCTION = "w:instr"
    LINK = "r:link"
    HREF = "r:href"
    SPACE = "xml:space"
    # The .rels attributes and DrawingML's typeface carry no prefix at all.
    TYPEFACE = "typeface"
    PANOSE = "panose"
    SCRIPT = "script"
    TARGET = "Target"
    TARGET_MODE = "TargetMode"
    RELATIONSHIP_TYPE = "Type"
    PART_NAME = "PartName"

    def of(self, value: str | int) -> Attribute:
        """`Attr.WIDTH.of(9360)` reads as the attribute it sets."""
        return Attribute(self, str(value))


class Value(StrEnum):
    """Attribute values that are OOXML keywords rather than data.

    One home for these. `dxa` was spelled as a literal in three modules, which
    is how a keyword becomes a typo that matches nothing and reports zero.
    """

    DXA = "dxa"  # twentieths of a point, the unit every width here is in
    PERCENT = "pct"
    FIXED = "fixed"
    AUTO = "auto"
    EXACT = "exact"
    CLEAR = "clear"
    CENTER = "center"
    PRESERVE = "preserve"
    EXTERNAL = "External"
    HYPERLINK = "hyperlink"


@dataclass(frozen=True, slots=True)
class Attribute:
    """One attribute and its value, so `make` never takes a bag of strings."""

    name: Attr
    value: str


XML_SPACE_PRESERVE: Final = Attribute(Attr.SPACE, Value.PRESERVE)

_PARSER: Final = etree.XMLParser(
    resolve_entities=False,  # no entity expansion: no XXE, no billion laughs
    no_network=True,  # never fetch a DTD, schema or entity
    load_dtd=False,
    dtd_validation=False,
    attribute_defaults=False,
    huge_tree=False,  # keep depth and text size bounded
    recover=False,  # a malformed part raises rather than half-parsing
)

SafeElement = NewType("SafeElement", etree._Element)


class UnsafePart(ValueError):
    """A part that carries a construct this package refuses to parse."""


def qn(name: Tag | Attr) -> str:
    """Turn `w:tbl` into the Clark notation lxml matches on."""
    prefix, _, local = name.partition(":")
    return f"{{{NS[prefix]}}}{local}" if local else prefix


def make(
    tag: Tag,
    *attributes: Attribute,
    children: Sequence[SafeElement] = (),
    text: str | None = None,
) -> SafeElement:
    """Build an element. The other constructor of `SafeElement`.

    Constructed markup is safe for a different reason than parsed markup: it
    never went through a parser at all, so there is no entity, no DTD and no
    input to smuggle either through. That is why building produces a
    `SafeElement` while `etree.fromstring` does not.
    """
    element = etree.Element(qn(tag))
    for attribute in attributes:
        element.set(qn(attribute.name), attribute.value)
    element.extend(children)
    if text is not None:
        element.text = text
    return SafeElement(element)


def parse_part(data: bytes, part: str) -> SafeElement:
    """Parse one part of a .docx. The only parsing constructor of `SafeElement`."""
    try:
        root = etree.fromstring(data, parser=_PARSER)
    except etree.XMLSyntaxError as exc:
        raise UnsafePart(f"{part} is not well-formed: {exc}") from exc
    # lxml-stubs does not declare DocInfo.doctype, though lxml has carried it
    # since 2.0 and the calibration tests below depend on it. getattr keeps the
    # gap in the stub from becoming a cast that would also hide a real change.
    doctype = str(getattr(root.getroottree().docinfo, "doctype", "") or "")
    if doctype:
        raise UnsafePart(f"{part} carries a DOCTYPE, which this package refuses to parse")
    return SafeElement(root)


def serialize_part(root: SafeElement) -> bytes:
    """Serialize a part back to bytes, declaration included."""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def children(root: SafeElement, tag: Tag) -> list[SafeElement]:
    """Direct children with this tag."""
    return [SafeElement(e) for e in root.findall(qn(tag))]


def descendants(root: SafeElement, tag: Tag) -> list[SafeElement]:
    """Every descendant with this tag, at any depth.

    A child of a safe tree is safe: it came through the same parser. Restating
    that here keeps the annotation honest down the tree rather than forcing a
    cast at every call site.
    """
    return [SafeElement(e) for e in root.iter(qn(tag))]


def child(root: SafeElement, tag: Tag) -> SafeElement | None:
    found = root.find(qn(tag))
    return None if found is None else SafeElement(found)


def attribute(element: SafeElement, name: Attr) -> str | None:
    """One attribute's value, or None."""
    value = element.get(qn(name))
    return None if value is None else str(value)


def value_of(element: SafeElement | None) -> str | None:
    """The `w:val` of an element, or None if the element is absent."""
    return None if element is None else attribute(element, Attr.VAL)


def text_of(element: SafeElement) -> str:
    """All descendant text, whitespace collapsed."""
    return " ".join("".join(t for t in element.itertext() if isinstance(t, str)).split())
