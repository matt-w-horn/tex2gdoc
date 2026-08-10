"""Table sizing paths that the author's own papers never reached.

Both branches below are needed for correctness and neither had ever executed, so
they were indistinguishable from dead code. These exercise them directly.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from tex2gdoc.tex2gdoc import TEXT_WIDTH, column_widths, parse_ooxml, qn, resize_tables

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _cell(text: str, span: int = 1) -> str:
    grid = f'<w:gridSpan w:val="{span}" />' if span > 1 else ""
    return (
        f"<w:tc><w:tcPr>{grid}</w:tcPr>"
        f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>'
    )


def _table(rows: list[str], columns: int) -> ET.Element:
    grid = "".join('<w:gridCol w:w="100" />' for _ in range(columns))
    body = "".join(f"<w:tr>{r}</w:tr>" for r in rows)
    return parse_ooxml(
        f'<w:document xmlns:w="{W}"><w:body><w:tbl><w:tblPr /><w:tblGrid>{grid}'
        f"</w:tblGrid>{body}</w:tbl></w:body></w:document>",
        "test",
    )


def test_a_spanning_cell_takes_the_width_of_the_columns_it_covers() -> None:
    """`\\multicolumn` becomes w:gridSpan, and its cell must get the summed width.

    Without the span arithmetic the column index falls out of step, so every cell
    after the merge is sized against the wrong column.
    """
    doc = _table([_cell("wide", span=2) + _cell("narrow"), _cell("a") + _cell("b") + _cell("c")], 3)
    assert resize_tables(doc) == 1

    widths = [
        [
            int(tc.find(qn("w:tcPr")).find(qn("w:tcW")).get(qn("w:w")))
            for tc in tr.findall(qn("w:tc"))
        ]
        for tr in doc.iter(qn("w:tr"))
    ]
    spanning_row, plain_row = widths
    assert sum(spanning_row) == TEXT_WIDTH
    assert sum(plain_row) == TEXT_WIDTH
    # The merged cell covers the first two columns, so it is their sum.
    assert spanning_row[0] == plain_row[0] + plain_row[1]


def test_columns_that_cannot_all_fit_shrink_in_proportion() -> None:
    """When every column's longest word is too wide, the loss is shared.

    The row still has to add up to the text width exactly, or the table renders
    ragged.
    """
    long_word = "x" * 400
    rows = [[_FakeCell(long_word), _FakeCell(long_word), _FakeCell(long_word)]]
    widths = column_widths(rows, 3)
    assert sum(widths) == TEXT_WIDTH
    assert all(w > 0 for w in widths)
    # Shared, not dumped on one column.
    assert max(widths) - min(widths) <= 2


class _FakeCell(ET.Element):
    """An element carrying one long unbreakable word."""

    def __init__(self, text: str) -> None:
        super().__init__(qn("w:tc"))
        run = ET.SubElement(self, qn("w:t"))
        run.text = text
