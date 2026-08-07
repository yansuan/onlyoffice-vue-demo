#!/usr/bin/env bash
# Build and run onlyoffice-vue-demo (app only; Document Server is separate).
set -euo pipefail
cd "$(dirname "$0")/.."

IMAGE="${IMAGE:-onlyoffice-vue-demo:latest}"
NAME="${NAME:-onlyoffice-vue-demo}"
PORT="${PORT:-19102}"
VITE_DOCUMENT_SERVER_URL="${VITE_DOCUMENT_SERVER_URL:-http://192.168.93.128:19101}"
VITE_CALLBACK_BASE_URL="${VITE_CALLBACK_BASE_URL:-http://192.168.93.1:19102}"
VITE_WS_BASE_URL="${VITE_WS_BASE_URL:-ws://192.168.93.1:19102}"
VITE_ONLYOFFICE_JWT_SECRET="${VITE_ONLYOFFICE_JWT_SECRET:-+keng2vx4V2ei1k/2wAsbxjpNP/v6Ew7uhyaJ9hgOr4=}"
VITE_DOCUMENT_PATH="${VITE_DOCUMENT_PATH:-/files/demo.docx}"

echo "Building $IMAGE ..."
docker build -t "$IMAGE" .

docker rm -f "$NAME" >/dev/null 2>&1 || true

echo "Starting $NAME on host port $PORT -> container 4000 ..."
docker run -d \
  --name "$NAME" \
  -p "${PORT}:4000" \
  -e "VITE_DOCUMENT_SERVER_URL=$VITE_DOCUMENT_SERVER_URL" \
  -e "VITE_CALLBACK_BASE_URL=$VITE_CALLBACK_BASE_URL" \
  -e "VITE_WS_BASE_URL=$VITE_WS_BASE_URL" \
  -e "VITE_ONLYOFFICE_JWT_SECRET=$VITE_ONLYOFFICE_JWT_SECRET" \
  -e "VITE_DOCUMENT_PATH=$VITE_DOCUMENT_PATH" \
  "$IMAGE"

echo "OK. Open ${VITE_CALLBACK_BASE_URL}/"
docker logs --tail 20 "$NAME"
