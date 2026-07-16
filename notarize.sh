#!/bin/bash
# Sign, notarize and staple Timex.app.
#
# Editing anything inside the bundle breaks the seal, so this has to be re-run
# after every code change that is meant to ship. Source of truth for the code is
# the bundle itself (there is no separate project), so it signs in place.
set -euo pipefail

APP="${APP:-/Applications/Timex.app}"
IDENTITY="${TIMEX_SIGN_IDENTITY:-Developer ID Application: Kostiantyn Halynskyi (PDB9JGGX74)}"
PROFILE="${NOTARY_KEYCHAIN_PROFILE:-Corder}"
HERE="$(cd "$(dirname "$0")" && pwd)"
ENTITLEMENTS="$HERE/Timex.entitlements"
ZIP="${TMPDIR:-/tmp}/Timex-notarize.zip"

say() { printf '\n\033[1m==> %s\033[0m\n' "$1"; }

say "Cleaning junk that would be sealed into the signature"
find "$APP" -name '.DS_Store' -delete 2>/dev/null || true
# Stale caches: scripts run as __main__ so these are never loaded, and one of
# them is even from a different Python version.
find "$APP" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true

say "Dropping dead binaries from the zip"
# py2app copies PIL's dylibs into python313.zip as well as next to _imaging.so.
# Nothing can dlopen out of a zip, so those copies are dead weight — but the
# notary service still unpacks the archive and rejects them as unsigned.
if unzip -l "$APP/Contents/Resources/lib/python313.zip" 2>/dev/null | grep -q 'PIL/.dylibs/'; then
  zip -d "$APP/Contents/Resources/lib/python313.zip" 'PIL/.dylibs/*' >/dev/null
fi

say "Signing nested libraries (inside-out; the bundle must be signed last)"
find "$APP" -type f \( -name '*.so' -o -name '*.dylib' \) -print0 \
  | xargs -0 -n1 codesign --force --options runtime --timestamp --sign "$IDENTITY"

say "Signing embedded frameworks (their binaries carry no .dylib suffix)"
for fw in Python.framework/Versions/3.13 Tcl.framework/Versions/8.6 Tk.framework/Versions/8.6; do
  [ -d "$APP/Contents/Frameworks/$fw" ] &&
    codesign --force --options runtime --timestamp --sign "$IDENTITY" \
      "$APP/Contents/Frameworks/$fw"
done

say "Signing the embedded notifier"
codesign --force --options runtime --timestamp --sign "$IDENTITY" \
  "$APP/Contents/Resources/TimexNotify.app/Contents/MacOS/timex-notify"
codesign --force --options runtime --timestamp --sign "$IDENTITY" \
  "$APP/Contents/Resources/TimexNotify.app"

say "Signing executables"
for exe in python Timex TimexMenubar; do
  codesign --force --options runtime --timestamp \
    --entitlements "$ENTITLEMENTS" --sign "$IDENTITY" "$APP/Contents/MacOS/$exe"
done

say "Signing the bundle"
codesign --force --options runtime --timestamp \
  --entitlements "$ENTITLEMENTS" --sign "$IDENTITY" "$APP"

say "Verifying"
codesign --verify --deep --strict --verbose=2 "$APP"
codesign -dvv "$APP" 2>&1 | grep -E 'Authority|flags' || true

say "Submitting to Apple (5-15 min)"
rm -f "$ZIP"
ditto -c -k --keepParent "$APP" "$ZIP"
xcrun notarytool submit "$ZIP" --keychain-profile "$PROFILE" --wait

say "Stapling"
xcrun stapler staple "$APP"
xcrun stapler validate "$APP"

say "Gatekeeper"
spctl --assess --type execute --verbose=2 "$APP"

say "Done"
