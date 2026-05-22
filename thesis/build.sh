#!/usr/bin/env bash
# Build the thesis from the single markdown source of truth, using the
# University-of-Patras / CEID template (XeLaTeX-adapted: template-xe.cls).
#
#   1. pandoc converts ../gkinis_konstantinos.md with thesis.template.tex
#      into the intermediate main.tex
#   2. xelatex (run 3× for TOC / refs) produces main.pdf
set -e
export PATH="/Library/TeX/texbin:$PATH"
cd "$(dirname "$0")"

SRC=../gkinis_konstantinos.md

# Keep the content images (referenced as ./*.png in the markdown) in sync
# with the repo-root originals so xelatex finds them in this directory.
for img in Carbon_Accounting_Scopes.png system_overview_diagram.png Extreme_Programming_Loops.png; do
  [ -f "../$img" ] && cp -f "../$img" "$img"
done

pandoc "$SRC" \
  --template=thesis.template.tex \
  --top-level-division=chapter \
  --highlight-style=tango \
  -t latex \
  -o main.tex

xelatex -interaction=nonstopmode main.tex
xelatex -interaction=nonstopmode main.tex
xelatex -interaction=nonstopmode main.tex

echo "Done — thesis/main.pdf"
