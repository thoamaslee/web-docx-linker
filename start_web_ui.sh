#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$ROOT_DIR/web_ui.pid"
LOG_FILE="$ROOT_DIR/web_ui.log"
HOST="${WEB_DOCX_HOST:-127.0.0.1}"
PORT="${WEB_DOCX_PORT:-8765}"

if [[ -f "$PID_FILE" ]]; then
  OLD_PID="$(cat "$PID_FILE" || true)"
  if [[ -n "$OLD_PID" ]] && kill -0 "$OLD_PID" 2>/dev/null; then
    echo "이미 실행 중입니다: http://$HOST:$PORT"
    exit 0
  fi
  rm -f "$PID_FILE"
fi

if curl --silent --fail "http://$HOST:$PORT/" >/dev/null 2>&1; then
  echo "이미 실행 중입니다: http://$HOST:$PORT"
  exit 0
fi

if [[ -x "$ROOT_DIR/.venv/bin/python" ]]; then
  PYTHON="$ROOT_DIR/.venv/bin/python"
else
  PYTHON="python3"
fi

cd "$ROOT_DIR"
nohup "$PYTHON" "$ROOT_DIR/web_ui.py" --host "$HOST" --port "$PORT" > "$LOG_FILE" 2>&1 &
PID="$!"
echo "$PID" > "$PID_FILE"
sleep 1

if ! kill -0 "$PID" 2>/dev/null; then
  echo "웹 UI 시작에 실패했습니다. 로그를 확인하세요: $LOG_FILE"
  rm -f "$PID_FILE"
  exit 1
fi

echo "웹 UI가 실행 중입니다: http://$HOST:$PORT"
echo "로그: $LOG_FILE"
