#!/bin/bash
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
PORT=8787

echo "==> Layout Shortcuts local server"
echo "    http://127.0.0.1:${PORT}/src/taskpane.html"
echo "    Leave this window open while using PowerPoint."
echo "    Press Ctrl+C to stop."
echo

cd "$DIR"
python3 -m http.server "$PORT" --bind 127.0.0.1
