#!/bin/bash
# Removes Cube Explorer from this Mac. Workbooks are not touched: the numbers already
# written into sheets stay, they just can no longer be refreshed.
set -uo pipefail

APP="$HOME/Library/Application Support/CubeBridge"
LABEL="com.cubebridge.helper"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
MANIFEST="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/cubebridge-manifest.xml"
KEYCHAIN="$HOME/Library/Keychains/login.keychain-db"
CA_NAME="CubeBridge Local CA (localhost only)"

cat <<EOF
This will remove:
  - the helper and its login item
  - the Cube Explorer add-in from Excel
  - saved connections, and their passwords from the Keychain
  - trust for the localhost certificate (macOS may ask for your password)
  - $APP
EOF
if [ "${1:-}" != "--yes" ]; then
  read -r -p "Continue? [y/N] " answer
  case "$answer" in y | Y | yes | YES) ;; *) echo "Nothing was removed."; exit 0 ;; esac
fi

launchctl bootout "gui/$(id -u)/$LABEL" >/dev/null 2>&1
rm -f "$PLIST" "$MANIFEST"

if [ -f "$APP/config.json" ]; then
  /usr/bin/python3 - "$APP/config.json" <<'EOF' | while read -r id; do
import json, sys
try:
    for connection in json.load(open(sys.argv[1])).get("connections", []):
        print(connection["id"])
except Exception:
    pass
EOF
    security delete-generic-password -s CubeBridge -a "$id" >/dev/null 2>&1
  done
fi

while security delete-certificate -c "$CA_NAME" -t "$KEYCHAIN" >/dev/null 2>&1; do :; done
rm -rf "$APP"
echo "Removed. Quit and reopen Excel to clear the button from the ribbon."
