#!/usr/bin/env bash
# Start Sheetly (builds the app on first run). `./sheetly.sh --install-desktop` adds a menu entry for the current
# user, pointing at this checkout. (Opening a file given on the command line is not supported yet.)
set -e
here="$(dirname "$(readlink -f "$0")")"
if [ "$1" = "--install-desktop" ]; then
  apps="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
  mkdir -p "$apps"
  cat > "$apps/sheetly.desktop" <<DESKTOP
[Desktop Entry]
Name=Sheetly
Comment=Spreadsheet with Excel-compatible formulas and shortcuts
Exec="$here/sheetly.sh"
Terminal=false
Type=Application
Categories=Office;Spreadsheet;
DESKTOP
  command -v update-desktop-database >/dev/null && update-desktop-database "$apps" || true
  echo "Installed $apps/sheetly.desktop"
  exit 0
fi
cd "$here"
[ -d dist ] || npx vite build
exec npx electron . --no-sandbox "$@"
