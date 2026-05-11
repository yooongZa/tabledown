#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP_NAME="Tabledown"
APP="$ROOT/dist/$APP_NAME.app"
PKG="$ROOT/dist/$APP_NAME-mas.pkg"
ENTITLEMENTS="$ROOT/entitlements/mac-app-store.plist"
STAGE_DIR="$(mktemp -d)"

trap 'rm -rf "$STAGE_DIR"' EXIT

cd "$ROOT"

find_identity() {
  local pattern="$1"
  security find-identity -v | awk -F '"' -v pattern="$pattern" '$2 ~ pattern { print $2; exit }'
}

profile_app_id() {
  local profile="$1"
  local decoded
  decoded="$(mktemp)"
  security cms -D -i "$profile" > "$decoded" 2>/dev/null || {
    rm -f "$decoded"
    return 1
  }
  plutil -extract Entitlements.com.apple.application-identifier raw -o - "$decoded" 2>/dev/null \
    || plutil -extract Entitlements.application-identifier raw -o - "$decoded" 2>/dev/null
  rm -f "$decoded"
}

find_provision_profile() {
  local bundle_id="$1"
  local profile
  for profile in "$HOME/Library/Developer/Xcode/UserData/Provisioning Profiles"/*.{provisionprofile,mobileprovision}; do
    [[ -f "$profile" ]] || continue
    local app_id
    app_id="$(profile_app_id "$profile" 2>/dev/null || true)"
    if [[ "$app_id" == *".$bundle_id" ]]; then
      printf '%s\n' "$profile"
      return 0
    fi
  done
  return 1
}

sign_macho_files() {
  local app_path="$1"
  local identity="$2"
  local entitlements="$3"
  while IFS= read -r -d '' file_path; do
    if file "$file_path" | grep -q 'Mach-O'; then
      if [[ "$file_path" == "$app_path/Contents/MacOS/"* ]]; then
        /usr/bin/codesign --force --sign "$identity" --entitlements "$entitlements" "$file_path"
      else
        /usr/bin/codesign --force --sign "$identity" "$file_path"
      fi
    fi
  done < <(find "$app_path/Contents" -type f -print0)
}

APP_SIGN_IDENTITY="${APP_SIGN_IDENTITY:-$(find_identity '^Apple Distribution:')}"
INSTALLER_SIGN_IDENTITY="${INSTALLER_SIGN_IDENTITY:-$(find_identity '^(3rd Party Mac Developer Installer|Mac Installer Distribution):')}"

if [[ -z "$APP_SIGN_IDENTITY" ]]; then
  echo "Missing Apple Distribution signing identity for Mac App Store app signing." >&2
  exit 1
fi

if [[ -z "$INSTALLER_SIGN_IDENTITY" ]]; then
  echo "Missing Mac Installer Distribution signing identity for Mac App Store .pkg signing." >&2
  echo "Create/download a Mac Installer Distribution certificate, then rerun this script." >&2
  exit 1
fi

rm -rf "$ROOT/build" "$APP" "$PKG"
COPYFILE_DISABLE=1 .venv/bin/python setup.py py2app

BUNDLE_ID="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleIdentifier' "$APP/Contents/Info.plist")"
PROFILE="${MAS_PROVISION_PROFILE:-$(find_provision_profile "$BUNDLE_ID" || true)}"

if [[ -z "$PROFILE" ]]; then
  echo "Missing Mac App Store provisioning profile for bundle id: $BUNDLE_ID" >&2
  echo "Create a Mac App Store Connect provisioning profile and set MAS_PROVISION_PROFILE if needed." >&2
  exit 1
fi

APP_STAGE="$STAGE_DIR/$APP_NAME.app"
ditto --noextattr --noacl "$APP" "$APP_STAGE"
cp "$PROFILE" "$APP_STAGE/Contents/embedded.provisionprofile"
xattr -cr "$APP_STAGE"
chmod -R u+rwX,go+rX "$APP_STAGE"

sign_macho_files "$APP_STAGE" "$APP_SIGN_IDENTITY" "$ENTITLEMENTS"
/usr/bin/codesign \
  --force \
  --sign "$APP_SIGN_IDENTITY" \
  --entitlements "$ENTITLEMENTS" \
  "$APP_STAGE"

/usr/bin/codesign --verify --deep --strict --verbose=2 "$APP_STAGE"
/usr/bin/codesign -d --entitlements :- "$APP_STAGE"

productbuild \
  --component "$APP_STAGE" /Applications \
  --sign "$INSTALLER_SIGN_IDENTITY" \
  "$PKG"

pkgutil --check-signature "$PKG"

if [[ -n "${APP_STORE_USERNAME:-}" && -n "${APP_STORE_PASSWORD:-}" ]]; then
  if [[ "${APP_STORE_UPLOAD:-0}" == "1" ]]; then
    xcrun altool \
      --upload-package "$PKG" \
      -u "$APP_STORE_USERNAME" \
      -p "$APP_STORE_PASSWORD" \
      --output-format json
  else
    xcrun altool \
      --validate-app "$PKG" \
      -u "$APP_STORE_USERNAME" \
      -p "$APP_STORE_PASSWORD" \
      --output-format json
  fi
fi

du -sh "$PKG"
