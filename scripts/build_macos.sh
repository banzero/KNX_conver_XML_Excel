#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"
export PYINSTALLER_CONFIG_DIR="$ROOT_DIR/.pyinstaller"
mkdir -p "$PYINSTALLER_CONFIG_DIR"

if ! command -v pyinstaller >/dev/null 2>&1; then
  python3 -m venv .venv-build
  source .venv-build/bin/activate
  python -m pip install --upgrade pip
  python -m pip install pyinstaller
fi

pyinstaller --clean --noconfirm --onefile --name knx-web-tool knx_web_tool.py

rm -rf package
mkdir -p package
cp dist/knx-web-tool package/
cp README_web_tool.md package/README_web_tool.md

rm -f knx-web-tool-macos.zip
(
  cd package
  zip -r ../knx-web-tool-macos.zip .
)

echo "Created: $ROOT_DIR/knx-web-tool-macos.zip"
