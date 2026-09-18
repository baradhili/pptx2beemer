#!/bin/bash
# Convert every pptx file in this script's directory to LaTeX beamer
# with the pptx2beemer pipeline (d2t), then render the generated TeX
# to a PDF with lualatex.
#
# Usage:
#   ./convert.sh [d2t options]
#
# Examples:
#   ./convert.sh        convert all pptx files here and render PDFs
#   ./convert.sh -d     ... with debug output (debug/ directory)
#
# Options are passed on to d2t (note: -p is unnecessary, the PDF step
# is built in and best-effort — a failed render only logs a warning).
# Output files (.tex, .xml, .svg/.pdf shape art, .pdf, .log, debug/)
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

# d2t replaces spaces in the basename with underscores for all outputs
tex_for() {
    local base
    base="$(basename "$1" .pptx)"
    echo "$SCRIPT_DIR/${base// /_}.tex"
}

render_pdf() {
    local tex="$1" base log
    base="${tex%.tex}"
    log="$base.lualatex.log"
    if ! command -v lualatex >/dev/null 2>&1; then
        echo "warning: lualatex not found, skipping PDF rendering"
        return 0
    fi
    echo "rendering PDF for $(basename "$base") ..."
    if (cd "$SCRIPT_DIR" && lualatex -interaction=nonstopmode "$(basename "$tex")" >"$log" 2>&1); then
        echo "writing pdf => $base.pdf"
    else
        echo "warning: PDF rendering failed, see $log"
        tail -n 15 "$log" || true
    fi
}

for f in "${pptx_files[@]}"; do
    echo ""
    echo "=== Converting $(basename "$f") ==="
    base="$(basename "$f" .pptx)"
    stem="${base// /_}"
    # remove stale pipeline outputs (shape art, extraction dir) of previous runs
    rm -f "$SCRIPT_DIR/$stem"-svg-*.svg "$SCRIPT_DIR/$stem"-svg-*.pdf
    rm -rf "$SCRIPT_DIR/$stem.pptx.tmp"
    "$D2T" "$@" "$f"
    tex="$(tex_for "$f")"
    if [ -f "$tex" ]; then
        render_pdf "$tex"
    else
        echo "warning: $tex not generated, skipping PDF rendering"
    fi
done
