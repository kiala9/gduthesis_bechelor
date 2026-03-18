#!/usr/bin/env sh
set -eu

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$ROOT"

if ! command -v latexmk >/dev/null 2>&1; then
  echo "[ERROR] latexmk not found in PATH." >&2
  echo "Install TeX Live or MiKTeX and ensure latexmk is available." >&2
  exit 1
fi

if ! command -v xelatex >/dev/null 2>&1; then
  echo "[ERROR] xelatex not found in PATH." >&2
  echo "Install TeX Live or MiKTeX and ensure xelatex is available." >&2
  exit 1
fi

if ! command -v bibtex >/dev/null 2>&1; then
  echo "[ERROR] bibtex not found in PATH." >&2
  echo "Install TeX Live or MiKTeX and ensure bibtex is available." >&2
  exit 1
fi

echo "[INFO] Building PDF with XeLaTeX..."
latexmk -xelatex -bibtex -interaction=nonstopmode main.tex
echo "[OK] Generated main.pdf"
