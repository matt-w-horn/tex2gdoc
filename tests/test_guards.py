"""Guards against defects that no output metric would reveal."""

from __future__ import annotations

import pytest

from tex2gdoc.tex2gdoc import REF_RE, parse_ooxml

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
    """Both attacks need a DOCTYPE, so refusing one removes the class."""
    with pytest.raises(ValueError, match="DTD"):
        parse_ooxml(f'<?xml version="1.0"?>{payload}<w:document {W}/>', "doc")


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
