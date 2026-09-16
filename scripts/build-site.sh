#!/usr/bin/env bash
set -euo pipefail

mapfile -d '' workbooks < <(
  find . -maxdepth 1 -type f \( -iname '*.xlsx' -o -iname '*.xls' -o -iname '*.xlsm' -o -iname '*.xlsb' \) -print0
)

if [[ ${#workbooks[@]} -ne 1 ]]; then
  echo "Expected exactly one Excel workbook in the repository root; found ${#workbooks[@]}." >&2
  exit 1
fi

build_dir="_site"
mkdir -p "$build_dir"
find "$build_dir" -mindepth 1 -maxdepth 1 -delete

cp index.html styles.css app.js "$build_dir/"
cp "${workbooks[0]}" "$build_dir/workbook.xlsx"

python3 - "${workbooks[0]}" "$build_dir/workbook.json" <<'PY'
import json
import sys
from pathlib import Path

source = Path(sys.argv[1])
destination = Path(sys.argv[2])
destination.write_text(
    json.dumps({"filename": source.name}, ensure_ascii=False),
    encoding="utf-8",
)
PY

touch "$build_dir/.nojekyll"
echo "Built read-only viewer for: ${workbooks[0]}"
