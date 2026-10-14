#!/bin/bash
# Cube Explorer, one-command install or update:
#
#   curl -fsSL https://raw.githubusercontent.com/danaenayati/excel-ssas-mac/main/get.sh | bash
#
# It puts the project in ~/CubeExplorer (or $CUBE_EXPLORER_DIR) and runs its install.sh,
# which needs nothing installed beforehand. Run the same command again to update.
#
# While the repository is private, the plain download above needs a GitHub login:
# this script then uses the GitHub command-line tool (`gh`) if it is signed in, or a
# git clone, which asks for your GitHub credentials.
set -euo pipefail

SLUG="danaenayati/excel-ssas-mac"
BRANCH="${CUBE_EXPLORER_BRANCH:-main}"
DEST="${CUBE_EXPLORER_DIR:-$HOME/CubeExplorer}"

say() { printf '==> %s\n' "$1"; }
fail() { printf '\nStopped: %s\n' "$1" >&2; exit 1; }

[ "$(uname -s)" = "Darwin" ] || fail "Cube Explorer is for Excel on a Mac."

# /usr/bin/git exists on every Mac, but only works once the Command Line Tools are in.
have_git() { command -v git >/dev/null 2>&1 && xcode-select -p >/dev/null 2>&1; }

if [ -d "$DEST/.git" ] && have_git; then
  say "Updating $DEST"
  git -C "$DEST" pull --ff-only --quiet || fail "could not update $DEST (local changes?)."
elif [ -e "$DEST" ] && [ ! -f "$DEST/install.sh" ]; then
  fail "$DEST exists and is not Cube Explorer. Move it away, or set CUBE_EXPLORER_DIR."
else
  tmp="$(mktemp -d)"
  trap 'rm -rf "$tmp"' EXIT
  say "Downloading Cube Explorer"
  # While the repository is private this download is refused; the next ways sign in.
  if curl -fsSL "https://codeload.github.com/$SLUG/tar.gz/refs/heads/$BRANCH" -o "$tmp/source.tar.gz" 2>/dev/null; then
    mkdir -p "$tmp/source"
    tar -xzf "$tmp/source.tar.gz" -C "$tmp/source" --strip-components 1
  elif command -v gh >/dev/null 2>&1 && gh auth status >/dev/null 2>&1; then
    gh repo clone "$SLUG" "$tmp/source" -- --quiet --branch "$BRANCH" || fail "gh could not download $SLUG."
  elif have_git; then
    git clone --quiet --branch "$BRANCH" "https://github.com/$SLUG.git" "$tmp/source" \
      || fail "git could not download $SLUG. Is your GitHub account allowed to see it?"
  else
    fail "the download needs a GitHub login while the project is private. Download it as a ZIP from https://github.com/$SLUG, unzip it, and run install.sh inside it."
  fi
  [ -f "$tmp/source/install.sh" ] || fail "the download does not look like Cube Explorer."
  # Keep nothing from an older copy but replace it whole: it holds no settings.
  rm -rf "$DEST"
  mkdir -p "$(dirname "$DEST")"
  mv "$tmp/source" "$DEST"
fi

say "Installing"
# Under `curl | bash` this script's input is the download; give the installer the terminal.
if [ ! -t 0 ] && (exec </dev/tty) 2>/dev/null; then
  bash "$DEST/install.sh" </dev/tty
else
  bash "$DEST/install.sh"
fi
