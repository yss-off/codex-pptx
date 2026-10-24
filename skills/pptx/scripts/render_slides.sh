#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo "Usage: $0 <presentation.pptx> [render-root]" >&2
  exit 2
fi

if [[ ! -f "$1" ]]; then
  echo "Presentation not found: $1" >&2
  exit 1
fi
input=$(realpath -- "$1")

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
render_root=${2:-"$(dirname -- "$input")/.pptx-renders"}
mkdir -p -- "$render_root"
render_root=$(realpath -- "$render_root")
run_dir=$(mktemp -d "$render_root/render-XXXXXX")

python "$script_dir/office/soffice.py" --headless --convert-to pdf --outdir "$run_dir" "$input" >/dev/null
base_name=$(basename -- "${input%.pptx}")
pdf_path="$run_dir/$base_name.pdf"

if [[ ! -f "$pdf_path" ]]; then
  echo "PDF conversion did not produce: $pdf_path" >&2
  exit 1
fi

pdftoppm -jpeg -r 150 "$pdf_path" "$run_dir/slide" >/dev/null

printf 'render_dir=%s\n' "$run_dir"
find "$run_dir" -maxdepth 1 -type f -name 'slide-*.jpg' -print | sort -V
