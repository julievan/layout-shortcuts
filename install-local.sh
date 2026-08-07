#!/bin/bash
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml"

cp "$DIR/manifest-local.xml" "$WEF"
cp "$DIR/manifest-local.xml" "$HOME/Downloads/manifest.xml"

echo "Installed local manifest to:"
echo "  $WEF"
echo
echo "Next:"
echo "  1. Run: $DIR/start-local.sh"
echo "  2. Quit PowerPoint (Cmd+Q), reopen"
echo "  3. Click Home -> Shortcuts once"
