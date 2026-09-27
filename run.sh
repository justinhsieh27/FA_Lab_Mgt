#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$SCRIPT_DIR/fa-lab.pid"
LOG_DIR="$SCRIPT_DIR/logs"
LOG_FILE="$LOG_DIR/fa-lab.log"
PORT="${PORT:-3000}"

cd "$SCRIPT_DIR"
mkdir -p "$LOG_DIR"

if [[ -f "$PID_FILE" ]]; then
  EXISTING_PID="$(<"$PID_FILE")"
  if kill -0 "$EXISTING_PID" 2>/dev/null; then
    echo "FA Lab Case Management is already running (PID $EXISTING_PID)."
    echo "Open: http://localhost:$PORT"
    exit 0
  fi
  rm -f "$PID_FILE"
fi

if command -v lsof >/dev/null 2>&1 && lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  echo "Port $PORT is already in use. Stop the process using that port or choose another port:" >&2
  echo "  PORT=3001 ./run.sh" >&2
  exit 1
fi

nohup env PORT="$PORT" npm run start >>"$LOG_FILE" 2>&1 < /dev/null &
SERVER_PID=$!
echo "$SERVER_PID" > "$PID_FILE"

sleep 1
if ! kill -0 "$SERVER_PID" 2>/dev/null; then
  echo "FA Lab Case Management did not start. Check: $LOG_FILE" >&2
  rm -f "$PID_FILE"
  exit 1
fi

echo "FA Lab Case Management started in the background (PID $SERVER_PID)."
echo "Open: http://localhost:$PORT"
echo "Log: $LOG_FILE"
