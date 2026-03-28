#!/usr/bin/env bash
set -euo pipefail

PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="${1:-.venv}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "Python interpreter not found: $PYTHON_BIN" >&2
  exit 1
fi

"$PYTHON_BIN" -m venv "$VENV_DIR"
# shellcheck disable=SC1090
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r "$SCRIPT_DIR/requirements.txt"

echo

echo "Virtual environment created in: $VENV_DIR"
echo "Activate it with: source $VENV_DIR/bin/activate"
echo "Then run for example:"
echo "  python $SCRIPT_DIR/md_template_pptx_builder.py --help"
