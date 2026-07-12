#!/usr/bin/env bash
# Build a Developer ID-signed, notarized Tabledown.dmg + notarized zip for
# GitHub Release distribution (separate from the Mac App Store .pkg).
#
# Prerequisites:
#   - "Developer ID Application" signing identity in the keychain
#   - a notarytool keychain profile holding App Store Connect credentials:
#       xcrun notarytool store-credentials "tabledown-notary" \
#         --apple-id <APPLE_ID> --team-id 495S4FVMCB --password <APP_SPECIFIC_PASSWORD>
#     (or an API key: --key <AuthKey.p8> --key-id <KEY_ID> --issuer <ISSUER_ID>)
#
# Usage:
#   NOTARY_PROFILE=tabledown-notary bash scripts/build_dmg.sh
# 또는 keychain 프로파일 없이 ASC API 키로 직접 (헤드리스 환경 — keychain 이
# GUI 잠금해제를 요구해 store-credentials 가 안 되는 경우):
#   NOTARY_KEY=~/.appstoreconnect/private_keys/AuthKey_XXXX.p8 \
#   NOTARY_KEY_ID=XXXX NOTARY_ISSUER=<ISSUER_ID> bash scripts/build_dmg.sh
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP_NAME="Tabledown"
DIST="$ROOT/dist"
APP="$DIST/$APP_NAME.app"
DMG="$DIST/$APP_NAME.dmg"
ZIP="$DIST/$APP_NAME-notarized.zip"
NOTARY_PROFILE="${NOTARY_PROFILE:-tabledown-notary}"
# ASC API 키 3종이 모두 주어지면 keychain 프로파일 대신 키 직접 인증.
if [[ -n "${NOTARY_KEY:-}" && -n "${NOTARY_KEY_ID:-}" && -n "${NOTARY_ISSUER:-}" ]]; then
  NOTARY_AUTH=(--key "$NOTARY_KEY" --key-id "$NOTARY_KEY_ID" --issuer "$NOTARY_ISSUER")
else
  NOTARY_AUTH=(--keychain-profile "$NOTARY_PROFILE")
fi
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT

cd "$ROOT"

SIGN_ID="${APP_SIGN_IDENTITY:-$(security find-identity -v | awk -F'"' '/Developer ID Application:/{print $2; exit}')}"
if [[ -z "$SIGN_ID" ]]; then
  echo "Missing 'Developer ID Application' signing identity." >&2
  exit 1
fi
echo "Signing identity: $SIGN_ID"

# 1. Fresh py2app build
rm -rf "$ROOT/build" "$APP" "$DMG" "$ZIP"
COPYFILE_DISABLE=1 .venv/bin/python setup.py py2app

# 2. Stage outside the iCloud-synced tree. iCloud reattaches a
#    com.apple.FinderInfo xattr to the bundle, which makes codesign fail with
#    "resource fork, Finder information, or similar detritus not allowed".
APP_STAGE="$STAGE/$APP_NAME.app"
ditto --noextattr --noacl "$APP" "$APP_STAGE"
xattr -cr "$APP_STAGE"

# 3. Sign nested Mach-O files first, then the bundle, with hardened runtime
#    and a secure timestamp (both required for notarization).
while IFS= read -r -d '' file_path; do
  if file "$file_path" | grep -q 'Mach-O'; then
    /usr/bin/codesign --force --options runtime --timestamp --sign "$SIGN_ID" "$file_path"
  fi
done < <(find "$APP_STAGE/Contents" -type f -print0)
/usr/bin/codesign --force --options runtime --timestamp --sign "$SIGN_ID" "$APP_STAGE"
/usr/bin/codesign --verify --deep --strict --verbose=2 "$APP_STAGE"

# 4. Notarize the app: submit the zipped app and wait for Apple's verdict.
NOTARIZE_ZIP="$STAGE/notarize.zip"
ditto -c -k --keepParent --norsrc --noextattr --noqtn "$APP_STAGE" "$NOTARIZE_ZIP"
xcrun notarytool submit "$NOTARIZE_ZIP" "${NOTARY_AUTH[@]}" --wait

# 5. Staple the ticket to the app, then copy the stapled app back to dist and
#    build the distribution zip from it (zips can't be stapled — the app
#    inside carries the ticket).
xcrun stapler staple "$APP_STAGE"
ditto --noextattr --noacl "$APP_STAGE" "$APP"
ditto -c -k --keepParent --norsrc --noextattr --noqtn "$APP_STAGE" "$ZIP"

# 6. Build the DMG from the stapled app, then notarize and staple the DMG
#    itself. The DMG is a separate artifact: notarizing the app alone is not
#    enough — `stapler staple <dmg>` needs a ticket issued for the DMG.
hdiutil create -volname "$APP_NAME" -srcfolder "$APP_STAGE" -ov -format UDZO "$DMG"
xcrun notarytool submit "$DMG" "${NOTARY_AUTH[@]}" --wait
xcrun stapler staple "$DMG"

# 6. Verify Gatekeeper acceptance and emit checksums for the release notes.
spctl -a -t open --context context:primary-signature -v "$DMG" || true
shasum -a 256 "$DMG" "$ZIP"
du -sh "$DMG" "$ZIP"
echo "Built and notarized: $DMG, $ZIP"
