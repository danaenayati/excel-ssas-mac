#!/bin/bash
# Installs Cube Explorer for Excel on this Mac. Safe to run again: it updates in place.
#
#   ./install.sh               install or update
#   ./install.sh --rebuild     also rebuild the Excel panel from source
#   ./install.sh --skip-trust  do everything except the step that asks for your password
#
# Nothing has to be installed first. The two tools the install needs are used from the
# Mac if present, and otherwise downloaded once into the folder below, each checked
# against a fixed SHA-256 checksum: uv (it installs a private Python 3.12 and the
# helper's libraries) and Node.js (it builds the Excel panel).
#
# What it changes on this Mac (everything is per-user; nothing needs sudo):
#   ~/Library/Application Support/CubeBridge/     helper, panel files, certificate, settings,
#                                                 logs, and the downloaded tools (tools/)
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
TOOLS="$APP/tools"

# Pinned versions and the checksums their publishers list for them.
UV_VERSION="0.12.5"
UV_SHA_ARM64="5bb0e5fe008a773c3dbcb97ff79cd89e1241464fe9d2f986d52ad8f1b037bd62"
UV_SHA_X64="b3b2137477cf96c9686ebfb71524614cec780c673fd73e59bce099aef02e70e8"
NODE_VERSION="24.21.0"
NODE_SHA_ARM64="bed7eea5325e1108f32ce5228ddd6a5f0f08a499ee42aa7442aea583702f6057"
NODE_SHA_X64="1462cb3b3046b815cf8ea436d3da450ec1a9f11dac7e5a46b0ada5305d7e8097"

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

[ "$(uname -s)" = "Darwin" ] || fail "Cube Explorer is for Excel on a Mac."
[ -d "/Applications/Microsoft Excel.app" ] || echo "Note: Microsoft Excel was not found in /Applications."
case "$(uname -m)" in
  arm64) ARCH="arm64" ;;
  x86_64) ARCH="x64" ;;
  *) fail "this Mac's processor ($(uname -m)) is not supported." ;;
esac

# Download a .tar.gz, check it against its pinned SHA-256, and unpack it into $TOOLS/<name>.
fetch_tool() {
  local name="$1" url="$2" sha="$3" tmp
  tmp="$(mktemp -d)"
  echo "Downloading $name (once)..."
  curl --fail --silent --show-error --location --retry 3 --output "$tmp/download.tar.gz" "$url" \
    || { rm -rf "$tmp"; fail "could not download $name from $url. Check the internet connection."; }
  echo "$sha  $tmp/download.tar.gz" | shasum -a 256 --check --status \
    || { rm -rf "$tmp"; fail "the download of $name does not match its published checksum. Nothing was installed."; }
  mkdir -p "$tmp/unpacked"
  tar -xzf "$tmp/download.tar.gz" -C "$tmp/unpacked" --strip-components 1
  rm -rf "${TOOLS:?}/$name"
  mkdir -p "$TOOLS"
  mv "$tmp/unpacked" "$TOOLS/$name"
  rm -rf "$tmp"
}

mkdir -p "$APP"
chmod 700 "$APP"

# uv: the Mac's own, else ours.
if command -v uv >/dev/null 2>&1; then
  UV="$(command -v uv)"
else
  UV="$TOOLS/uv/uv"
  if [ ! -x "$UV" ] || [ "$("$UV" --version 2>/dev/null | awk '{print $2}')" != "$UV_VERSION" ]; then
    if [ "$ARCH" = "arm64" ]; then target="aarch64-apple-darwin"; sha="$UV_SHA_ARM64"
    else target="x86_64-apple-darwin"; sha="$UV_SHA_X64"; fi
    fetch_tool uv "https://github.com/astral-sh/uv/releases/download/$UV_VERSION/uv-$target.tar.gz" "$sha"
  fi
  # The Python uv downloads stays in our folder too.
  export UV_PYTHON_INSTALL_DIR="$TOOLS/python"
fi

step "Installing the helper"
"$UV" venv --quiet --python 3.12 --allow-existing "$APP/venv"
# gssapi and krb5 (for Windows Authentication / Kerberos) use the Mac's own Kerberos
# library. They come as ready-made wheels for both kinds of Mac; --only-binary makes
# sure nothing is ever compiled here, even for a new version that ships source first.
"$UV" pip install --quiet --python "$PYTHON" --only-binary gssapi --only-binary krb5 \
  --reinstall-package cubebridge "$REPO/helper"
"$PYTHON" -m cubebridge --version >/dev/null || fail "the helper did not install correctly."

step "Installing the Excel panel"
# functions.js, commands.js: a build from before the cube formulas or the right-click
# commands existed is built again.
if [ -n "$REBUILD" ] || [ ! -f "$REPO/addin/dist/index.html" ] || [ ! -f "$REPO/addin/dist/functions.js" ] \
  || [ ! -f "$REPO/addin/dist/commands.js" ]; then
  # Node.js: the Mac's own when it is new enough for the build (20.19+ or 22.12+), else ours.
  if command -v node >/dev/null 2>&1 && command -v npm >/dev/null 2>&1 \
    && node -e 'const [a,b]=process.versions.node.split(".").map(Number); process.exit(a>22||(a===22&&b>=12)||(a===20&&b>=19)?0:1)'; then
    NODE_BIN="$(dirname "$(command -v node)")"
  else
    NODE_BIN="$TOOLS/node/bin"
    if [ ! -x "$NODE_BIN/node" ] || [ "$("$NODE_BIN/node" --version)" != "v$NODE_VERSION" ]; then
      if [ "$ARCH" = "arm64" ]; then sha="$NODE_SHA_ARM64"; else sha="$NODE_SHA_X64"; fi
      fetch_tool node "https://nodejs.org/dist/v$NODE_VERSION/node-v$NODE_VERSION-darwin-$ARCH.tar.gz" "$sha"
    fi
  fi
  echo "Building the panel..."
  (cd "$REPO/addin" && PATH="$NODE_BIN:$PATH" npm ci --no-audit --no-fund --loglevel=error \
    && PATH="$NODE_BIN:$PATH" npm run build >/dev/null) \
    || fail "the panel could not be built. Run ./install.sh --rebuild again; if it keeps failing, see the messages above."
fi
[ -f "$REPO/addin/dist/index.html" ] || fail "the panel build did not produce addin/dist/index.html."
[ -f "$REPO/addin/dist/functions.js" ] || fail "the panel build did not produce addin/dist/functions.js."
[ -f "$REPO/addin/dist/commands.js" ] || fail "the panel build did not produce addin/dist/commands.js."
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

cat <<EOF

Done. Cube Explorer is installed.

Next:
  1. Quit Excel completely if it is open (Cmd+Q), then open it again.
  2. In Excel: Home tab > Add-ins (far right of the ribbon) > Cube Explorer.
     A "Cube Explorer" button then appears at the far right of the Home tab.
  3. Try it with the built-in "Sample cube" first. It needs no server.
  4. To check your company server, connect the VPN and run:  "$REPO/cube-probe"

To remove it later:  "$REPO/uninstall.sh"
EOF
if pgrep -xq "Microsoft Excel"; then
  echo
  echo "Excel is open now: save your work, quit it (Cmd+Q) and open it again to see the add-in."
fi
