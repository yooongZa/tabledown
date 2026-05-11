#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP="$ROOT/dist/Tabledown.app"
ZIP="$ROOT/dist/Tabledown.zip"
VERIFY_DIR="$(mktemp -d)"

trap 'rm -rf "$VERIFY_DIR"' EXIT

cd "$ROOT"

rm -rf "$ROOT/build" "$APP" "$ZIP"
COPYFILE_DISABLE=1 .venv/bin/python setup.py py2app

# macOS can attach FinderInfo xattrs to bundles in synced folders. Codesign
# treats those as resource-fork detritus, so remove only that disallowed xattr.
find "$APP" -xattrname com.apple.FinderInfo -exec xattr -d com.apple.FinderInfo {} \;

/usr/bin/codesign --verify --deep --strict --verbose=2 "$APP"

ditto -c -k --keepParent --norsrc --noextattr --noqtn --noacl \
  --zlibCompressionLevel 9 "$APP" "$ZIP"
ditto -x -k "$ZIP" "$VERIFY_DIR"
/usr/bin/codesign --verify --deep --strict --verbose=2 "$VERIFY_DIR/Tabledown.app"

du -sh "$APP" "$ZIP"
