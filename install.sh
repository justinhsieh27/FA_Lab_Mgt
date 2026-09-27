#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if ! command -v node >/dev/null 2>&1 || ! command -v npm >/dev/null 2>&1; then
  echo "Node.js 22+ and npm are required. Please install Node.js before continuing." >&2
  exit 1
fi

echo "Installing FA Lab Case Management dependencies…"
npm install
echo "Installation complete. Start the system with: ./run.sh"
