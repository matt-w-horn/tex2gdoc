"""Table sizing paths that the author's own papers never reached.

Both branches below are needed for correctness and neither had ever executed, so
they were indistinguishable from dead code. These exercise them directly.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from tex2gdoc.ooxml import Attr, SafeElement, Tag, attribute, child, children, descendants, make
from tex2gdoc.tex2gdoc import (
    TEXT_WIDTH,
    column_widths,
    parse_ooxml,
    resize_tables,
    restyle_docx,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _cell(text: str, span: int = 1) -> str:
    grid = f'<w:gridSpan w:val="{span}" />' if span > 1 else ""
    return (
        f"<w:tc><w:tcPr>{grid}</w:tcPr>"
        f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>'
    )


def _table(rows: list[str], columns: int) -> SafeElement:
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

    def cell_width(cell: SafeElement) -> int:
        properties = child(cell, Tag.CELL_PROPERTIES)
        assert properties is not None
        width = child(properties, Tag.CELL_WIDTH)
        assert width is not None
        return int(attribute(width, Attr.WIDTH) or 0)

    widths = [
        [cell_width(tc) for tc in children(row, Tag.TABLE_CELL)]
        for row in descendants(doc, Tag.TABLE_ROW)
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
    rows = [[_fake_cell(long_word), _fake_cell(long_word), _fake_cell(long_word)]]
    widths = column_widths(rows, 3)
    assert sum(widths) == TEXT_WIDTH
    assert all(w > 0 for w in widths)
    # Shared, not dumped on one column.
    assert max(widths) - min(widths) <= 2


def _fake_cell(text: str) -> SafeElement:
    """A cell carrying one long unbreakable word."""
    return make(Tag.TABLE_CELL, children=[make(Tag.TEXT, text=text)])


def test_restyle_drops_an_empty_comments_part_and_its_declarations(tmp_path: Path) -> None:
    """The pass that removes a part must leave the package internally consistent.

    Three things have to move together: the part, its Content_Types override and
    its relationship. A dangling override makes Word refuse the file outright and
    makes Docs fail silently, which is the same no-error symptom this change set
    was chasing in the first place. `restyle_docx` also writes from the zip
    listing it read, which still names the deleted part, and that KeyError only
    appeared on a real conversion because nothing in the calibration suite
    reaches this code.
    """
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Override PartName="/word/comments.xml" ContentType="application/comments+xml" />'
        '<Override PartName="/word/document.xml" ContentType="application/document+xml" />'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://x/comments" Target="comments.xml" />'
        '<Relationship Id="rId2" Type="http://x/styles" Target="styles.xml" />'
        "</Relationships>"
    )
    path = tmp_path / "case.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("word/document.xml", f'<w:document xmlns:w="{W}"><w:body /></w:document>')
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/comments.xml", f'<w:comments xmlns:w="{W}" />')
        z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W}" />')

    assert restyle_docx(path).empty_comments_dropped

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        types = z.read("[Content_Types].xml").decode()
        written_rels = z.read("word/_rels/document.xml.rels").decode()
    assert "word/comments.xml" not in names
    assert {"word/document.xml", "word/styles.xml"} <= set(names)
    assert "/word/comments.xml" not in types, "dangling Content_Types override"
    assert "/word/document.xml" in types, "unrelated override removed"
    assert "comments.xml" not in written_rels, "dangling relationship"
    assert "styles.xml" in written_rels, "unrelated relationship removed"


def test_restyle_keeps_a_comments_part_that_holds_comments(tmp_path: Path) -> None:
    """Only an empty one goes. A real comment must survive the sweep."""
    path = tmp_path / "case.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", f'<w:document xmlns:w="{W}"><w:body /></w:document>')
        z.writestr(
            "word/comments.xml",
            f'<w:comments xmlns:w="{W}"><w:comment w:id="1"><w:p /></w:comment></w:comments>',
        )
    assert not restyle_docx(path).empty_comments_dropped
    with zipfile.ZipFile(path) as z:
        assert "word/comments.xml" in z.namelist()
