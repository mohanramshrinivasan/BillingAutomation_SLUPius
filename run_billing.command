#!/bin/zsh
cd "$(dirname "$0")" || exit 1

if command -v python3 >/dev/null 2>&1; then
    python3 billing_prepare.py
else
    echo "Python 3 is not installed."
fi

echo ""
echo "Press Enter to close..."
read