#!/usr/bin/env sh
set -eu

ROOT=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
cd "$ROOT"

if ! command -v git >/dev/null 2>&1; then
  echo "[ERROR] git not found in PATH." >&2
  exit 1
fi

if [ ! -d .git ]; then
  echo "[ERROR] Current directory is not a git repository." >&2
  exit 1
fi

if [ "$#" -ne 3 ]; then
  echo "Usage: sh scripts/make_release_patch.sh <old-ref> <new-ref> <output.patch>" >&2
  echo "Example: sh scripts/make_release_patch.sh v1.0.0 v1.0.1 release_v1.0.1.patch" >&2
  exit 1
fi

OLD_REF=$1
NEW_REF=$2
OUTPUT=$3

echo "[INFO] Generating patch from $OLD_REF to $NEW_REF ..."
git diff --binary "$OLD_REF" "$NEW_REF" > "$OUTPUT"
echo "[OK] Patch written to $OUTPUT"
