#!/usr/bin/env bash
# Capture the README's demo screenshots from the running add-in.
#
# Requires Screen Recording permission for the terminal you run this from:
# System Settings → Privacy & Security → Screen & System Audio Recording.
#
# Usage: open scripts/big.xlsx in Excel with the Crunched pane showing the
# state you want, then run:
#
#     ./scripts/capture-demo-screenshots.sh 01-workbook-overview
#
set -euo pipefail

name="${1:?usage: capture-demo-screenshots.sh <name-without-extension>}"
root="$(cd "$(dirname "$0")/.." && pwd)"
out="$root/docs/screenshots/$name.png"
mkdir -p "$(dirname "$out")"

# -l takes a CGWindowID; ask macOS for Excel's frontmost document window.
window_id="$(
  osascript -e 'tell application "System Events" to tell process "Microsoft Excel" to get value of attribute "AXIdentifier" of window 1' 2>/dev/null || true
)"

if [[ -n "${window_id}" && "${window_id}" =~ ^[0-9]+$ ]]; then
  screencapture -x -o -l"${window_id}" "$out"
else
  echo "Click the Excel window when the crosshair appears."
  screencapture -x -o -w "$out"
fi

echo "wrote ${out#"$root"/}"
