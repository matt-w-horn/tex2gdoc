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

Wrote examples/sample.docx (43 KB). Work files kept in examples/sample-work/

Verification (run against the .docx itself):
  [PASS] figures embedded           1 image(s) in the .docx, 1 figure(s) in the source
  [PASS] tables present             1 table(s) in the .docx, 1 tabular(s) in the source
  [PASS] table layout               1 table(s) at 9360 dxa, 6 cell(s) sized in dxa
  [PASS] column widths sum          every grid sums to 9360 dxa
  [PASS] figures centered           figure and caption styles are centered
  [PASS] page geometry              margins set to 1440 dxa
  [PASS] code highlighted           27 coloured code run(s)
  [PASS] table contents survived    1 table(s) probed, all contents found
  [PASS] captions survived          2 caption(s) probed, all found
  [PASS] code listings readable     1 listing(s) probed, spacing intact
  [PASS] cross-references resolved  no raw label keys in the text
  [PASS] citations resolved         1 key(s) cited, all resolved
  [PASS] bibliography rendered      reference section found, 138 chars of entries (floor 40)
  [PASS] math converted             3 Word equation(s), 4 inline-math span(s) in the source (ratio 0.75; exact parity is not expected)
  [PASS] sections present           3 top-level heading(s), 2 \section(s) in the source
  [PASS] body text volume           779 chars extracted, 567 in the source (ratio 1.37)

Notes:
  - Flattened 1 nested tabular cell(s); pandoc drops those tables otherwise.
  - Unstarred 1 table*/figure* environment(s); pandoc drops the caption of a starred float.
  - Resolved 2 table/figure/theorem cross-reference(s) to numbers.
  - Converted 1 alltt listing(s) to verbatim to keep their spacing; math escapes became literal characters.
```

The command exits nonzero if any check reports FAIL, so a bad `.docx` never
passes as a good one.

## Install

You need `pandoc`, a TeX engine, and `pdftocairo` from Poppler.

```bash
brew install pandoc poppler tectonic
```

```bash
pip install -e '.[dev]'
```

`tectonic` is preferred over `pdflatex`. It fetches the packages a figure needs,
so no full TeX distribution is required, and it is XeTeX-based, which matches how
most modern papers are built. Pass `--tex-engine pdflatex` to use the other one.

## Use

```bash
tex2gdoc paper.tex
```

It picks up `references.bib` beside the source, writes `paper.docx`, and keeps
its intermediate files in `paper-work/` for inspection.

| Command | What it does |
|---|---|
| `make check` | ruff and pytest. Needs no pandoc and no TeX. |
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
  silent. This is the known weak point.

## Styling

Only styles that carry layout are set: figure centring, the code font, table
widths, one inch margins, 1.15 line spacing. Pandoc's own definitions of
`Heading1` through `Heading3`, `Title` and `Caption` are left alone, so Word and
Google Docs map them onto their native heading styles and their style menus keep
working on the imported document.

## How it checks itself

The output is measured against the source, and the two measurements are
deliberately independent. Measuring the source with pandoc's own reader would let
a pandoc defect cancel out on both sides, and every check would go green on a
broken document. All four losses in the table above were found by exactly this
comparison.

There are 18 checks and every one is calibrated. `--self-test` builds input
designed to break each check and asserts that it reports FAIL. A check added
without a calibration case fails the coverage gate. None of it needs pandoc or a
TeX engine, so all of it runs in CI.

Two properties this buys, both from real defects:

- A check that stops firing is caught, not only a check that fails. Several
  checks appear only when the source has something to check, so a parsing bug can
  make a check vanish rather than fail. The known-good case asserts the exact set
  of check names.
- Every edit to the `.docx` and every check on it works on a parsed XML tree
  rather than on serialized text, so a pandoc upgrade that changes tag spacing
  cannot silently zero a count.

## Layout

| Path | What |
|---|---|
| `src/tex2gdoc/tex2gdoc.py` | the converter |
| `src/tex2gdoc/verification.py` | the 18 checks |
| `src/tex2gdoc/_calibration.py` | one case per check, proving it can fail |
| `src/tex2gdoc/baseline.py` | structural fingerprint of a `.docx`, and the comparison |
| `examples/sample.tex` | the paper in the example above |

The generated `.docx` is never committed. It is derived from the `.tex` and the
`.bib`, so a stale copy in the tree is worse than no copy.
