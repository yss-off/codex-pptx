#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 <script.js> [args...]" >&2
  exit 2
fi

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
skill_dir=$(cd -- "$script_dir/.." && pwd)

if [[ -d "$skill_dir/node_modules/pptxgenjs" ]]; then
  node_path="$skill_dir/node_modules"
elif command -v npm >/dev/null 2>&1; then
  node_path=$(npm root -g)
else
  echo "PptxGenJS is unavailable: npm was not found." >&2
  exit 1
fi

if [[ ! -d "$node_path/pptxgenjs" ]]; then
  echo "PptxGenJS is unavailable. Install it locally in the skill or globally with npm." >&2
  exit 1
fi

NODE_PATH="$node_path${NODE_PATH:+:$NODE_PATH}" exec node "$@"
