# One command per thing worth doing.
#
#   make check       what CI runs. Needs no pandoc and no TeX engine.
#   make docx TEX=…  regenerate a review copy, and refuse to hand over a bad one.
#
# The regression and baseline targets need a paper, which lives outside this
# repo, so they take TEX= rather than assuming one.

VENV := .venv
PY   := $(VENV)/bin/python
TEX  ?=

.PHONY: help check lint format typecheck test selftest docx regression baseline clean

help:
	@echo "make check                 ruff + mypy + pytest; no pandoc or TeX needed"
	@echo "make format                autoformat and autofix in place"
	@echo "make selftest              prove every check fails on input built to break it"
	@echo "make docx TEX=paper.tex    regenerate and verify a .docx"
	@echo "make regression TEX=…      convert and compare against the recorded baseline"
	@echo "make baseline TEX=…        record a new baseline beside that paper"

check: lint typecheck test

lint:
	$(VENV)/bin/ruff check .
	$(VENV)/bin/ruff format --check .

# Autoformat and autofix in place. `lint` only reports; this one edits.
format:
	$(VENV)/bin/ruff check --fix .
	$(VENV)/bin/ruff format .

typecheck:
	$(VENV)/bin/mypy

test:
	$(VENV)/bin/pytest -q

selftest:
	$(VENV)/bin/tex2gdoc --self-test

# tex2gdoc exits nonzero when any check reports FAIL, so this target already
# refuses to leave a bad .docx looking like a good one.
docx:
	@test -n "$(TEX)" || { echo "usage: make docx TEX=path/to/paper.tex" >&2; exit 2; }
	$(VENV)/bin/tex2gdoc "$(TEX)"

regression:
	@test -n "$(TEX)" || { echo "usage: make regression TEX=path/to/paper.tex" >&2; exit 2; }
	TEX2GDOC_REGRESSION_TEX="$(TEX)" $(VENV)/bin/pytest -q -m regression

baseline:
	@test -n "$(TEX)" || { echo "usage: make baseline TEX=path/to/paper.tex" >&2; exit 2; }
	$(PY) scripts/record_baseline.py "$(TEX)"

clean:
	rm -rf .pytest_cache .ruff_cache .mypy_cache
	find . -name __pycache__ -type d -prune -exec rm -rf {} +
