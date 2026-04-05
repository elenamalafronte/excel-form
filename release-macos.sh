#!/usr/bin/env bash
set -euo pipefail

APP_NAME="ExcelForm"

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

PYTHON_BIN="${PYTHON_BIN:-python3}"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "Python executable not found: $PYTHON_BIN"
  exit 1
fi

echo "Using Python: $($PYTHON_BIN --version 2>&1)"

# Isolated environment for deterministic build behavior.
VENV_DIR=".venv-macos-build"
if [ ! -d "$VENV_DIR" ]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

python -m pip install --upgrade pip
python -m pip install pyinstaller customtkinter openpyxl

mkdir -p installer-output

echo "Building macOS app bundle..."
PYINSTALLER_ARGS=(
  --noconfirm
  --clean
  --windowed
  --name "$APP_NAME"
)

if [ -f "Heat number summary.xlsm" ]; then
  PYINSTALLER_ARGS+=(--add-data "Heat number summary.xlsm:.")
else
  echo "Note: Heat number summary.xlsm not found in repo checkout; building without bundled sample workbook."
fi

PYINSTALLER_ARGS+=(main.py)
pyinstaller "${PYINSTALLER_ARGS[@]}"

APP_BUNDLE_PATH="dist/${APP_NAME}.app"
if [ ! -d "$APP_BUNDLE_PATH" ]; then
  echo "Build failed: ${APP_BUNDLE_PATH} was not created."
  exit 1
fi

ZIP_PATH="installer-output/${APP_NAME}-macOS.zip"
rm -f "$ZIP_PATH"

echo "Packaging app bundle as zip..."
ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE_PATH" "$ZIP_PATH"

echo "Done. Send this file to macOS clients:"
echo "  $ZIP_PATH"
