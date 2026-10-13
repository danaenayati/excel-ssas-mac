#!/bin/bash
# Installs Cube Explorer for Excel on this Mac. Safe to run again: it updates in place.
#
#   ./install.sh               install or update
#   ./install.sh --rebuild     also rebuild the Excel panel from source (needs Node.js)
#   ./install.sh --skip-trust  do everything except the step that asks for your password
#
# What it changes on this Mac (everything is per-user; nothing needs sudo):
#   ~/Library/Application Support/CubeBridge/     helper, panel files, certificate, settings, logs
#   ~/Library/LaunchAgents/com.cubebridge.helper.plist    starts the helper at login
#   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/cubebridge-manifest.xml
#   login keychain: trust for one certificate that is valid for https://localhost only
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
APP="$HOME/Library/Application Support/CubeBridge"
PORT=47613
LABEL="com.cubebridge.helper"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
KEYCHAIN="$HOME/Library/Keychains/login.keychain-db"
CA_NAME="CubeBridge Local CA (localhost only)"
PYTHON="$APP/venv/bin/python"

REBUILD=""
SKIP_TRUST=""
for arg in "$@"; do
  case "$arg" in
    --rebuild) REBUILD="yes" ;;
    --skip-trust) SKIP_TRUST="yes" ;;
    *) echo "Unknown option: $arg" >&2; exit 2 ;;
  esac
done

step() { printf '\n==> %s\n' "$1"; }
fail() { printf '\nInstall stopped: %s\n' "$1" >&2; exit 1; }

command -v uv >/dev/null 2>&1 || fail "the 'uv' tool is missing. Install it with:  brew install uv"
[ -d "/Applications/Microsoft Excel.app" ] || echo "Note: Microsoft Excel was not found in /Applications."

step "Installing the helper"
mkdir -p "$APP"
chmod 700 "$APP"
uv venv --quiet --python 3.12 --allow-existing "$APP/venv"
uv pip install --quiet --python "$PYTHON" --reinstall-package cubebridge "$REPO/helper"
"$PYTHON" -m cubebridge --version >/dev/null || fail "the helper did not install correctly."

step "Installing the Excel panel"
if [ -n "$REBUILD" ] || [ ! -f "$REPO/addin/dist/index.html" ]; then
  command -v npm >/dev/null 2>&1 || fail "the panel is not built yet and Node.js (npm) is missing."
  (cd "$REPO/addin" && npm ci --no-audit --no-fund && npm run build)
fi
[ -f "$REPO/addin/dist/index.html" ] || fail "the panel build did not produce addin/dist/index.html."
rm -rf "$APP/pane.new"
cp -R "$REPO/addin/dist" "$APP/pane.new"
rm -rf "$APP/pane"
mv "$APP/pane.new" "$APP/pane"

step "Preparing the certificate for https://localhost"
"$PYTHON" -m cubebridge setup --dir "$APP" >/dev/null
TRUSTED="yes"
if security verify-cert -c "$APP/certs/localhost.pem" -p ssl -s localhost >/dev/null 2>&1; then
  echo "Already trusted."
elif [ -n "$SKIP_TRUST" ]; then
  TRUSTED=""
  echo "Skipped. Excel cannot load the panel until you run ./install.sh without --skip-trust."
else
  # Remove certificates from earlier installs so they do not pile up.
  while security delete-certificate -c "$CA_NAME" -t "$KEYCHAIN" >/dev/null 2>&1; do :; done
  echo "macOS will now ask for your Mac password (or Touch ID)."
  echo "This lets Excel trust the helper on this Mac only. The certificate is valid for"
  echo "https://localhost and nothing else, and the key that created it has been destroyed."
  security add-trusted-cert -r trustRoot -p ssl -k "$KEYCHAIN" "$APP/certs/ca.pem" \
    || fail "the certificate was not trusted (the password prompt was cancelled?). Run ./install.sh again."
  security verify-cert -c "$APP/certs/localhost.pem" -p ssl -s localhost >/dev/null 2>&1 \
    || fail "macOS still does not trust the certificate."
  echo "Trusted."
fi

step "Starting the helper (and at every login)"
mkdir -p "$HOME/Library/LaunchAgents" "$APP/logs"
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$PYTHON</string>
    <string>-m</string>
    <string>cubebridge</string>
    <string>serve</string>
    <string>--dir</string>
    <string>$APP</string>
    <string>--port</string>
    <string>$PORT</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>ThrottleInterval</key><integer>10</integer>
  <key>ProcessType</key><string>Interactive</string>
  <key>StandardOutPath</key><string>$APP/logs/launchd.log</string>
  <key>StandardErrorPath</key><string>$APP/logs/launchd.log</string>
</dict>
</plist>
EOF
launchctl bootout "gui/$(id -u)/$LABEL" >/dev/null 2>&1 || true
# bootout returns before the old copy is fully gone; starting too early fails.
for _ in $(seq 1 20); do
  launchctl print "gui/$(id -u)/$LABEL" >/dev/null 2>&1 || break
  sleep 0.25
done
launchctl bootstrap "gui/$(id -u)" "$PLIST" || fail "launchd refused to start the helper."

healthy=""
for _ in $(seq 1 30); do
  if curl --silent --fail --noproxy '*' --cacert "$APP/certs/ca.pem" --output /dev/null "https://localhost:$PORT/"; then
    healthy="yes"
    break
  fi
  sleep 0.5
done
[ -n "$healthy" ] || fail "the helper is not answering on https://localhost:$PORT. See $APP/logs/helper.log"
echo "Helper is running."

step "Adding Cube Explorer to Excel"
mkdir -p "$WEF"
cp "$REPO/addin/manifest/manifest.xml" "$WEF/cubebridge-manifest.xml"

if [ -z "$TRUSTED" ]; then
  printf '\nAlmost done. Run ./install.sh once more (without --skip-trust) to finish.\n'
  exit 0
fi

cat <<'EOF'

Done.

Next:
  1. Quit Excel completely if it is open (Cmd+Q), then open it again.
  2. In Excel: Home tab > Add-ins (far right of the ribbon) > Cube Explorer.
     A "Cube Explorer" button then appears at the far right of the Home tab.
  3. Try it with the built-in "Sample cube" first. It needs no server.
  4. To check your company server, connect the VPN and run:  ./cube-probe
EOF
