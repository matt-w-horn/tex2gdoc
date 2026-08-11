"""Guards against defects that no output metric would reveal."""

from __future__ import annotations

import zipfile
from collections.abc import Callable
from pathlib import Path

import pytest

from tex2gdoc.ooxml import UnsafePart
from tex2gdoc.tex2gdoc import REF_RE, SourceFacts, parse_ooxml
from tex2gdoc.verification import Status, verify

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def test_parse_ooxml_accepts_an_ordinary_part() -> None:
    root = parse_ooxml(f'<?xml version="1.0"?><w:document {W}><w:body/></w:document>', "doc")
    assert root.tag.endswith("}document")


@pytest.mark.parametrize(
    "payload",
    [
        pytest.param(
            '<!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;">]>',
            id="entity-expansion",
        ),
        pytest.param(
            '<!DOCTYPE d [<!ENTITY x SYSTEM "file:///etc/passwd">]>',
            id="external-entity",
        ),
    ],
)
def test_parse_ooxml_refuses_a_dtd(payload: str) -> None:
    """Both attacks need a DOCTYPE, so refusing one removes the class.

    The refusal is now the parser's own `docinfo.doctype` rather than a scan of
    the first 8 KB, so a declaration pushed past that window by a long prolog
    comment no longer slips through. That case is the third parameter below.
    """
    with pytest.raises(UnsafePart, match="DOCTYPE"):
        parse_ooxml(f'<?xml version="1.0"?>{payload}<w:document {W}/>', "doc")


def test_parse_ooxml_refuses_a_dtd_past_the_old_scan_window() -> None:
    """The exact bypass the previous 8 KB regex scan documented and allowed."""
    padding = f"<!--{'x' * 9000}-->"
    payload = '<!DOCTYPE d [<!ENTITY x SYSTEM "file:///etc/passwd">]>'
    with pytest.raises(UnsafePart, match="DOCTYPE"):
        parse_ooxml(f'<?xml version="1.0"?>{padding}{payload}<w:document {W}/>', "doc")


def test_ref_re_never_matches_a_section_label() -> None:
    """Pandoc numbers sections correctly on its own; this resolver must not.

    Widening REF_RE to cover `sec:` looks like removing a special case. It would
    overwrite pandoc's hierarchical numbers such as "3.2" with this script's flat
    per-label counter, and no check in verify() would notice, because the
    cross-reference check only looks for raw keys that were left behind.
    """
    assert REF_RE.search(r"\ref{tab:map}")
    assert REF_RE.search(r"\ref{fig:phase}")
    assert REF_RE.search(r"\ref{thm:clamp}")
    assert not REF_RE.search(r"\ref{sec:model}")
    assert not REF_RE.search(r"\ref{app:map}")
    assert not REF_RE.search(r"\ref{eq:kernel}")


@pytest.mark.parametrize(
    ("label", "build"),
    [
        ("not a zip", lambda p: p.write_bytes(b"this is not a zip file at all")),
        ("empty file", lambda p: p.write_bytes(b"")),
    ],
)
def test_verify_reports_an_unreadable_container_instead_of_raising(
    tmp_path: Path, label: str, build: Callable[[Path], object]
) -> None:
    """`verify` is documented as reading a file that came back from a co-author.

    A malformed XML part already reported `document parses: FAIL`; a file that
    was not a .docx at all reached the caller as a zipfile traceback, which is
    the one input most likely to arrive by accident.
    """
    docx = tmp_path / f"{label.replace(' ', '-')}.docx"
    build(docx)
    checks = verify(docx, SourceFacts(), "")
    assert [c.name for c in checks] == ["output exists"] or checks[0].name == "document parses"
    assert checks[0].status is Status.FAIL


def test_verify_refuses_a_part_it_cannot_decode(tmp_path: Path) -> None:
    """Undecodable bytes must fail, not be repaired before inspection.

    `read_docx` used to decode with errors="replace" and hand the repaired text
    to the parser, so every check inspected a document Word would never render.
    Passing the bytes straight to the parser makes the file fail instead.
    """
    docx = tmp_path / "bad-bytes.docx"
    with zipfile.ZipFile(docx, "w") as z:
        z.writestr(
            "word/document.xml",
            f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}>'.encode()
            + b"<w:body><w:p><w:r><w:t>\xff\xfe</w:t></w:r></w:p></w:body></w:document>",
        )
    checks = verify(docx, SourceFacts(), "")
    assert checks[0].name == "document parses"
    assert checks[0].status is Status.FAIL
