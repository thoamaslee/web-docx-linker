#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$ROOT_DIR/web_ui.pid"

if [[ ! -f "$PID_FILE" ]]; then
  PORT="${WEB_DOCX_PORT:-8765}"
  PIDS="$(lsof -tiTCP:"$PORT" -sTCP:LISTEN 2>/dev/null || true)"
  if [[ -z "$PIDS" ]]; then
    echo "실행 중인 웹 UI를 찾지 못했습니다."
    exit 0
  fi
  kill $PIDS
  echo "웹 UI를 종료했습니다."
  exit 0
fi

PID="$(cat "$PID_FILE" || true)"
if [[ -n "$PID" ]] && kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  echo "웹 UI를 종료했습니다."
else
  echo "웹 UI 프로세스를 찾지 못했습니다."
fi

rm -f "$PID_FILE"
