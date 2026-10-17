#!/usr/bin/env bash
# Start Sheetly (builds the app on first run). `./sheetly.sh --install-desktop` adds a menu entry for the current
# user, pointing at this checkout, and makes Sheetly an "Open with" choice for .xlsx files.
# ./sheetly.sh book.xlsx opens a workbook; with Sheetly already running it opens in a new window there.
set -e
here="$(dirname "$(readlink -f "$0")")"
if [ "$1" = "--install-desktop" ]; then
  apps="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
  mkdir -p "$apps"
  cat > "$apps/sheetly.desktop" <<DESKTOP
[Desktop Entry]
Name=Sheetly
Comment=Spreadsheet with Excel-compatible formulas and shortcuts
Exec="$here/sheetly.sh" %F
Terminal=false
Type=Application
Categories=Office;Spreadsheet;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
DESKTOP
  command -v update-desktop-database >/dev/null && update-desktop-database "$apps" || true
  echo "Installed $apps/sheetly.desktop"
  exit 0
fi
# Workbook paths are made absolute before changing to the app directory.
args=()
for a in "$@"; do if [ -e "$a" ]; then args+=("$(readlink -f "$a")"); else args+=("$a"); fi; done
cd "$here"
[ -d dist ] || npx vite build
exec npx electron . --no-sandbox "${args[@]}"
