#!/usr/bin/env sh
set -eu

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
SCRIPTS="$ROOT/scripts"
TOOLS="$ROOT/tools"
EXPORT_TEX="$ROOT/pandoc_export.tex"
OUTPUT_DOCX="$ROOT/thesis.docx"

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

if ! command -v pandoc >/dev/null 2>&1; then
  echo "[ERROR] pandoc not found in PATH." >&2
  echo "Install Pandoc and ensure it is available." >&2
  exit 1
fi

if command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD=python3
elif command -v python >/dev/null 2>&1; then
  PYTHON_CMD=python
else
  echo "[ERROR] Python 3 not found in PATH." >&2
  echo "Install Python 3 and ensure python3 or python is available." >&2
  exit 1
fi

echo "[INFO] Refreshing LaTeX references and bibliography..."
latexmk -xelatex -bibtex -interaction=nonstopmode main.tex

echo "[INFO] Building pandoc export source..."
"$PYTHON_CMD" "$SCRIPTS/build_pandoc_export.py" "$ROOT" --output "$EXPORT_TEX"

echo "[INFO] Converting to Word..."
pandoc "$EXPORT_TEX" \
  --from=latex \
  --to=docx \
  --output="$OUTPUT_DOCX" \
  --reference-doc="$TOOLS/Gthesis_reference.docx" \
  --lua-filter="$SCRIPTS/pandoc_keyword_style.lua" \
  --lua-filter="$SCRIPTS/pandoc_bibliography_list.lua" \
  --lua-filter="$SCRIPTS/pandoc_caption_numbering.lua" \
  --resource-path="$ROOT:$ROOT/figures" \
  --top-level-division=chapter \
  --number-sections \
  --toc \
  -M suppress-bibliography=true

echo "[INFO] Post-processing Word document..."
"$PYTHON_CMD" "$SCRIPTS/fix_docx_bibliography.py" "$OUTPUT_DOCX"

echo "[OK] Generated thesis.docx"
