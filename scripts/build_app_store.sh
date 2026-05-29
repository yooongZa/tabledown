#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP_NAME="Tabledown"
APP="$ROOT/dist/$APP_NAME.app"
PKG="$ROOT/dist/$APP_NAME-mas.pkg"
ENTITLEMENTS="$ROOT/entitlements/mac-app-store.plist"
INHERIT_ENTITLEMENTS="$ROOT/entitlements/mac-app-store-inherit.plist"
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
  # PlistBuddy uses ':' path separators, so the key's own dots (e.g.
  # com.apple.application-identifier) are kept literal. `plutil -extract` treats
  # dots as path separators and fails to find the key.
  local app_id
  app_id="$(/usr/libexec/PlistBuddy -c 'Print :Entitlements:com.apple.application-identifier' "$decoded" 2>/dev/null \
    || /usr/libexec/PlistBuddy -c 'Print :Entitlements:application-identifier' "$decoded" 2>/dev/null)"
  rm -f "$decoded"
  [[ -n "$app_id" ]] && printf '%s\n' "$app_id"
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
  local entitlements="$3"          # full entitlements (with application-identifier)
  local inherit_entitlements="$4"  # sandbox + inherit only (no application-identifier)
  local main_exe
  main_exe="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleExecutable' "$app_path/Contents/Info.plist")"
  while IFS= read -r -d '' file_path; do
    if file "$file_path" | grep -q 'Mach-O'; then
      if [[ "$file_path" == "$app_path/Contents/MacOS/$main_exe" ]]; then
        # Main executable: full entitlements; the app bundle carries the
        # embedded provisioning profile that matches application-identifier.
        /usr/bin/codesign --force --sign "$identity" --entitlements "$entitlements" "$file_path"
      elif [[ "$file_path" == "$app_path/Contents/MacOS/"* ]]; then
        # Other nested executables (e.g. py2app's python interpreter) must NOT
        # carry application-identifier, or App Store Connect rejects the build
        # (error 90885: nested executable has app id but no provisioning
        # profile). Give them sandbox-inherit so they inherit the main app's
        # profile instead of requiring their own.
        /usr/bin/codesign --force --sign "$identity" --entitlements "$inherit_entitlements" "$file_path"
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

sign_macho_files "$APP_STAGE" "$APP_SIGN_IDENTITY" "$ENTITLEMENTS" "$INHERIT_ENTITLEMENTS"
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
