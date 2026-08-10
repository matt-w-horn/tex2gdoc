#!/usr/bin/env python3
"""Record a structural baseline for a regression target.

Runs the converter over a paper, measures the .docx it produces, and writes the
fingerprint plus every check status to JSON. The default destination sits beside
the paper, in its own repo's gitignored `.verify/` directory, because the
baseline describes that paper and must never be committed to this repo.

This script holds only generic counting logic and no paper content, so it is
safe to track here.

Usage:
    scripts/record_baseline.py paper.tex
    scripts/record_baseline.py paper.tex --out /somewhere/baseline.json
    scripts/record_baseline.py --docx existing.docx --out baseline.json
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tex2gdoc.baseline import measure_docx  # noqa: E402
from tex2gdoc.tex2gdoc import convert, default_tex_engine  # noqa: E402


def tool_version(name: str, *args: str) -> str:
    try:
        out = subprocess.run([name, *args], capture_output=True, text=True, timeout=30)
        return (out.stdout or out.stderr).strip().splitlines()[0]
    except Exception as exc:  # noqa: BLE001 - a missing tool is data, not a crash
        return f"unavailable ({type(exc).__name__})"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("tex", nargs="?", type=Path, help="the paper to convert and measure")
    ap.add_argument(
        "--docx", type=Path, help="measure an existing .docx instead of converting (checks omitted)"
    )
    ap.add_argument("--out", type=Path, help="where to write the JSON")
    ap.add_argument("--note", default="", help="free text recorded alongside the numbers")
    args = ap.parse_args()

    if not args.tex and not args.docx:
        ap.error("give a .tex to convert, or --docx to measure an existing file")

    engine = default_tex_engine()
    record: dict[str, object] = {
        "note": args.note,
        "engine": engine,
        "versions": {
            "pandoc": tool_version("pandoc", "--version"),
            "tex": tool_version(engine, "--version"),
            "pdftocairo": tool_version("pdftocairo", "-v"),
        },
    }

    if args.docx:
        docx = args.docx
        record["checks"] = None
        record["source"] = f"measured from an existing artifact: {docx.name}"
    else:
        tmp = Path(tempfile.mkdtemp(prefix="tex2gdoc-baseline-"))
        docx = tmp / "baseline.docx"
        # Same default the CLI uses. Passing None here would quietly record a
        # baseline with no bibliography, so every later run would look like a
        # regression against it.
        bib = args.tex.parent / "references.bib"
        bib = bib if bib.exists() else None
        record["bibliography"] = bib.name if bib else None
        checks, notes = convert(args.tex, docx, bib, 300, tmp / "work", engine=engine)
        record["checks"] = {c.name: c.status for c in checks}
        record["notes"] = notes
        record["source"] = f"converted from {args.tex.name}"
        failed = [c.name for c in checks if c.status == "FAIL"]
        if failed:
            print(f"WARNING: recording a baseline with failing checks: {failed}", file=sys.stderr)

    record["counts"] = measure_docx(docx)
    if not args.docx:
        # Keep the artifact next to its numbers. It is the thing a later run gets
        # diffed against by eye when a count moves and the reason is not obvious.
        keep = args.out.parent if args.out else args.tex.parent / ".verify"
        keep.mkdir(parents=True, exist_ok=True)
        shutil.copy(docx, keep / "tex2gdoc-baseline.docx")

    out = args.out or (args.tex.parent / ".verify" / "tex2gdoc-baseline.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"wrote {out}")
    for key, value in sorted(record["counts"].items()):  # type: ignore[union-attr]
        print(f"  {key:34} {value}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
