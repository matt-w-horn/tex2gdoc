# tex2gdoc

Converts a LaTeX paper to a `.docx` that opens cleanly in Microsoft Word or
Google Docs, then checks the file it produced.

Pandoc already converts LaTeX to `.docx` in one command. On a real paper it also
drops content, and it drops it silently: exit status 0, an empty log, and a
document that looks finished. Four losses show up again and again.

| What you wrote | What arrives without help |
|---|---|
| A `tabular` nested inside a table cell | the whole table, gone, leaving an empty shell |
| A caption on a `table*` or `figure*` | no caption |
| An `alltt` code listing | `structure Loop where` becomes `structureLoopwhere` |
| `\ref` to a table or figure | the raw key, `[tab:cases]` |

TikZ and PGFPlots figures do not survive at all, because pandoc runs no TeX
engine. This renders each one with a real TeX run, fixes the four losses above,
and then verifies the `.docx` against what the source claimed.

There is a fifth failure that is not a loss. Google Docs refused a 22-page paper
built from pandoc's defaults outright, with no error message of any kind, and
took the same paper once this converter had been through it. Which difference
Docs choked on is [not fully pinned](#styling).

## A worked example

`examples/sample.tex` is a short paper carrying every construct above.

```bash
tex2gdoc examples/sample.tex --bib examples/sample.bib
```

```
Source: 1 TikZ figure(s), 0 image figure(s), 1 table(s), 1 citation key(s)
Page metrics: columnwidth=229.5pt, normalsize=10pt, 5 size step(s) (after \maketitle; engine: tectonic)
  throughput: rendered (36 KB)
Styled: 1 table(s) set to the full text width, 1 listing(s) highlighted
Styles: 9 paragraph(s) moved onto a style Google Docs knows; 0 kept auto leading for tall maths
Fonts: 1 embedding directive(s) removed, 101 typeface name(s) dropped from the theme and font table (Aptos Display, 游ゴシック Light, and 45 more) so the reader's template supplies them

Wrote examples/sample.docx (43 KB). Work files kept in examples/sample-work/

Verification (run against the .docx itself):
  [PASS] figures embedded            1 image(s) in the .docx, 1 figure(s) in the source
  [PASS] tables present              1 table(s) in the .docx, 1 tabular(s) in the source
  [PASS] table layout                1 table(s) at 9360 dxa, 6 cell(s) sized in dxa
  [PASS] column widths sum           every grid sums to 9360 dxa
  [PASS] figures centered            figure and caption styles are centered
  [PASS] page geometry               margins set to 1440 dxa
  [PASS] code highlighted            27 coloured code run(s)
  [PASS] table contents survived     1 table(s) probed, all contents found
  [PASS] captions survived           2 caption(s) probed, all found
  [PASS] code listings readable      1 listing(s) probed, spacing intact
  [PASS] cross-references resolved   no raw label keys in the text
  [PASS] citations resolved          1 key(s) cited, all resolved
  [PASS] bibliography rendered       reference section found, 138 chars of entries (floor 40)
  [PASS] math converted              3 Word equation(s), 4 inline-math span(s) in the source (ratio 0.75; exact parity is not expected)
  [PASS] sections present            3 top-level heading(s), 2 \section(s) in the source
  [PASS] body text volume            779 chars extracted, 567 in the source (ratio 1.37)
  [PASS] body styles are native      running prose is on Normal, which Docs maps to its own body style
  [PASS] no external dependencies    nothing is fetched or embedded on open; 0 hyperlink(s), followed on click not on open
  [PASS] fonts left to the template  no typeface named; the theme is empty and the body inherits

Notes:
  - Flattened 1 nested tabular cell(s); pandoc drops those tables otherwise.
  - Unstarred 1 table*/figure* environment(s); pandoc drops the caption of a starred float.
  - Resolved 2 table/figure/theorem cross-reference(s) to numbers.
  - Converted 1 alltt listing(s) to verbatim to keep their spacing; math escapes became literal characters.
```

The command exits nonzero if any check reports FAIL, so a bad `.docx` never
passes as a good one.

## Install

```bash
brew install matt-w-horn/tap/tex2gdoc
```

That pulls in the three tools it shells out to: `pandoc` for the conversion,
`tectonic` for the figures pandoc cannot render, and `pdftocairo` from Poppler
to rasterize them. Nothing else needs setting up first.

`tectonic` is preferred over `pdflatex`. It fetches the packages a figure needs,
so no full TeX distribution is required, and it is XeTeX-based, which matches how
most modern papers are built. Pass `--tex-engine pdflatex` to use the other one.

### From source

For working on the tool rather than using it. Python 3.14, and the same three
binaries.

```bash
brew install pandoc poppler tectonic
```

```bash
pip install -e '.[dev]'
```

## Use

```bash
tex2gdoc paper.tex
```

It picks up `references.bib` beside the source, writes `paper.docx`, and keeps
its intermediate files in `paper-work/` for inspection.

| Command | What it does |
|---|---|
| `make check` | ruff, mypy and pytest. Needs no pandoc and no TeX. |
| `make format` | autoformats and autofixes in place; `make check` only reports |
| `make selftest` | proves every check fails on input built to break it |
| `make docx TEX=paper.tex` | regenerates and verifies |
| `make baseline TEX=paper.tex` | records a structural baseline beside that paper |
| `make regression TEX=paper.tex` | converts and compares against that baseline |

## What it does not do

- **Convert back.** LaTeX to `.docx` only. For the other direction, `pandoc
  file.docx -o file.tex` is usually enough, because Word documents do not carry
  the constructs that break going this way.
- **Render `\begin{cases}` correctly in Google Docs.** Math becomes native Word
  equations and Docs renders them, but a piecewise definition arrives as one run
  with its branches run together. `\mathbb{R}` comes through as an italic `R`.
  Both were confirmed on 2026-08-09 by opening a generated file in Docs. Word
  itself handles both.
- **Preview through LibreOffice.** It renders no Word equations at all, not even
  `$x$`, so it shows blanks where the math is.
- **Parse arbitrary LaTeX.** The source is read with regular expressions tuned to
  the constructs above. An unusual nesting will not match, and a non-match is
  silent. This is the known weak point, and it is confined to the LaTeX side:
  the `.docx` is read and written as a parsed tree, never matched as text.

## Styling

Only styles that carry layout are set: figure centring, the code font, table
widths, one inch margins, 1.15 line spacing. Pandoc's own definitions of
`Heading1` through `Heading3`, `Title` and `Caption` are left alone. Word and
Google Docs then map them onto their native heading styles, and their style
menus keep working on the imported document.

Two things this converter takes away, so the destination gets to decide them.
The first matters more than it sounds. A 22-page paper carrying pandoc's
untouched font handling crashed Google Docs' importer on 2026-08-10, with no
error message of any kind. The same paper, with the handling below, imported
cleanly. That build changed four things at once, so which one Docs choked on is
not yet pinned.

The theme names no typeface. Pandoc's reference document ships the Microsoft 365
theme, which names Aptos, Aptos Display, and a face per script for Japanese,
Korean, Chinese, Arabic and the rest: 94 names, almost none of which a reader
outside a current Office install has. All 94 are cleared, along with the PANOSE
metrics that would otherwise steer substitution back to something Aptos-shaped.
Font table entries nothing references any more go too, which is why the run
reports 101 names dropped and not 94. Two names survive on purpose:
`Courier New`, because no theme slot supplies monospace and a listing set in a
proportional face stops being a listing, and `Cambria Math`, which is the only
surviving hint about how the equations should look.

Paragraphs are moved onto styles Docs has. Its menu offers Normal text, Title,
Subtitle and Heading 1-6, and it turns every other style it meets into direct
formatting that the menu cannot reach, so a paper imported on pandoc's own
styles can have its headings restyled and nothing else. `BodyText`,
`FirstParagraph`, `Compact` and `Bibliography` become `Normal`, which takes a
22-page paper from 11% of its paragraphs reachable from the menu to 97%.

Nothing moves on the page, and the set is picked so that stays true. Those four
resolve to exactly what `Normal` resolves to in pandoc's own styles: two carry
the same spacing, two carry no paragraph properties at all, and `Compact`,
which differs only in having no space after, has that one property pinned back
on the paragraph so table cells do not gain 8pt each. `Abstract`, `BlockText`,
`Author`, the captions and `SourceCode` keep their own styles, because each
carries something no native style does: keep-with-next, the indents that make a
block quote a block quote, centring, the monospace block. Together they are 19
paragraphs out of 667.

Line spacing is pinned to an exact height. Under Word's default "multiple"
spacing every line is as tall as the tallest thing on it. A line carrying an
inline equation therefore grows, and a page of prose develops uneven leading
that LaTeX, with its fixed baseline, never shows. An exact height holds every
line to the same depth.
The paragraphs holding an equation taller than a line, n-ary operators with
limits, fractions, radicals and matrices, keep proportional spacing, because a
fixed height crops what overflows it: 32 of 667 paragraphs in that same paper.

## How it checks itself

The output is measured against the source, and the two measurements are
deliberately independent. Measuring the source with pandoc's own reader would let
a pandoc defect cancel out on both sides, and every check would go green on a
broken document. All four losses in the table above were found by exactly this
comparison.

There are 21 checks and every one is calibrated. `--self-test` builds input
designed to break each check and asserts that it reports FAIL. A check added
without a calibration case fails the coverage gate. None of it needs pandoc or a
TeX engine, so all of it runs in CI.

Eighteen of the 21 ask whether the content survived. The other three ask a
different question: what happens when you hand the file to someone else. It
must be self-contained, so opening it fetches nothing over the network and
embeds no font. It must name no typeface, so it takes the fonts of whatever
template it lands in. And its running prose must sit on `Normal`, because
Google Docs flattens any style it cannot map into direct formatting, which its
own style menu can no longer reach.

Three properties this buys, each one paid for by a defect that got through:

- A check that stops firing is caught too. Several checks appear only when the
  source has something for them to check. A parsing bug can therefore make a
  check vanish silently while every remaining check still reports PASS. The
  known-good case asserts the exact set of check names.
- Every edit to the `.docx`, and every check on it, works on a parsed XML tree.
  Nothing here matches serialized text. A pandoc upgrade that changes tag
  spacing cannot silently zero a count.
- One module, `ooxml.py`, owns the only XML parser, with entity expansion, DTD
  loading and network access switched off. Two mechanisms keep it that way:
  ruff bans importing every XML parser in the stdlib and in lxml anywhere else,
  and the tree type is a `NewType`, so passing a tree that did not come from
  the hardened parser is a type error at the call site. Both have limits worth
  knowing. Being a list of module names, the ban cannot cover a parser nobody
  has thought of, and the `NewType` constructor can be called deliberately,
  which this package does to relabel loosely-typed `.iter()` results.

## Layout

| Path | What |
|---|---|
| `src/tex2gdoc/tex2gdoc.py` | the converter |
| `src/tex2gdoc/ooxml.py` | the only XML parser, the element builder, and the tag and attribute names |
| `src/tex2gdoc/verification.py` | the 21 checks |
| `src/tex2gdoc/_calibration.py` | one case per check, proving it can fail |
| `src/tex2gdoc/baseline.py` | structural fingerprint of a `.docx`, and the comparison |
| `examples/sample.tex` | the paper in the example above |

The generated `.docx` is never committed. It is derived from the `.tex` and the
`.bib`, so a stale copy in the tree is worse than no copy.
