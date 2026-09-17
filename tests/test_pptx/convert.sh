#!/bin/bash
# Convert every pptx file in this script's directory to LaTeX beamer
# with the pptx2beemer pipeline (d2t).
#
# Usage:
#   ./convert.sh [d2t options]
#
# Examples:
#   ./convert.sh        convert all pptx files in this directory
#   ./convert.sh -d     ... with debug output (debug/ directory)
#   ./convert.sh -p     ... and compile a PDF with lualatex
#
# Options are passed on to d2t; see d2t -h for the full list.
# Output files (.tex, .xml, .csv template, debug/, .pdf with -p)
# are written next to the pptx files.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
D2T="$SCRIPT_DIR/../../d2t"

if [ ! -x "$D2T" ]; then
    echo "Error: d2t pipeline script not found at $D2T" >&2
    exit 1
fi

shopt -s nullglob
pptx_files=("$SCRIPT_DIR"/*.pptx)
shopt -u nullglob

if [ ${#pptx_files[@]} -eq 0 ]; then
    echo "No pptx files found in $SCRIPT_DIR" >&2
    exit 1
fi

for f in "${pptx_files[@]}"; do
    echo ""
    echo "=== Converting $(basename "$f") ==="
    "$D2T" "$@" "$f"
done
