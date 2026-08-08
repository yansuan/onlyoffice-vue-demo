#!/usr/bin/env bash
# Build and run onlyoffice-vue-demo, then deploy jsapi-executor into Document Server.
# Usage:
#   ./scripts/install.sh
#   bash scripts/install.sh
# Override any env before run, e.g.:
#   VITE_CALLBACK_BASE_URL=http://x.x.x.x:19102 ./scripts/install.sh
# Re-exec under bash when invoked as `sh scripts/install.sh` (dash has no pipefail).
if [ -z "${BASH_VERSION:-}" ]; then
  exec bash "$0" "$@"
fi
set -euo pipefail
cd "$(dirname "$0")/.."

IMAGE="${IMAGE:-onlyoffice-vue-demo:latest}"
NAME="${NAME:-onlyoffice-vue-demo}"
PORT="${PORT:-19102}"
VITE_DOCUMENT_SERVER_URL="${VITE_DOCUMENT_SERVER_URL:-http://121.41.239.31:19101}"
VITE_CALLBACK_BASE_URL="${VITE_CALLBACK_BASE_URL:-http://121.41.239.31:19102}"
VITE_WS_BASE_URL="${VITE_WS_BASE_URL:-ws://121.41.239.31:19102}"
VITE_ONLYOFFICE_JWT_SECRET="${VITE_ONLYOFFICE_JWT_SECRET:-+keng2vx4V2ei1k/2wAsbxjpNP/v6Ew7uhyaJ9hgOr4=}"
VITE_DOCUMENT_PATH="${VITE_DOCUMENT_PATH:-/files/demo.docx}"

# Document Server (plugin deploy)
CONTAINER="${CONTAINER:-documentserver}"
PLUGIN_SRC="plugins/jsapi-executor"
PLUGIN_DEST="/var/www/onlyoffice/documentserver/sdkjs-plugins"

echo "==> Building $IMAGE ..."
docker build -t "$IMAGE" .

docker rm -f "$NAME" >/dev/null 2>&1 || true

echo "==> Starting $NAME on host port $PORT -> container 4000 ..."
docker run -d \
  --name "$NAME" \
  -p "${PORT}:4000" \
  -e "VITE_DOCUMENT_SERVER_URL=$VITE_DOCUMENT_SERVER_URL" \
  -e "VITE_CALLBACK_BASE_URL=$VITE_CALLBACK_BASE_URL" \
  -e "VITE_WS_BASE_URL=$VITE_WS_BASE_URL" \
  -e "VITE_ONLYOFFICE_JWT_SECRET=$VITE_ONLYOFFICE_JWT_SECRET" \
  -e "VITE_DOCUMENT_PATH=$VITE_DOCUMENT_PATH" \
  "$IMAGE"

echo "==> App OK. Open ${VITE_CALLBACK_BASE_URL}/"
echo "    config: ${VITE_CALLBACK_BASE_URL}/config.js"
docker logs --tail 20 "$NAME"

# --- OnlyOffice plugin deploy (same as scripts/oo.sh) ---
if ! docker inspect "$CONTAINER" >/dev/null 2>&1; then
  echo "==> Skip plugin deploy: container '$CONTAINER' not found (see docs/docker.txt)." >&2
  exit 0
fi

if [[ ! -d "$PLUGIN_SRC" ]]; then
  echo "Plugin directory not found: $PLUGIN_SRC" >&2
  exit 1
fi

echo "==> Remove old plugin in $CONTAINER ..."
docker exec "$CONTAINER" rm -rf "${PLUGIN_DEST}/jsapi-executor"

echo "==> Copy $PLUGIN_SRC -> ${CONTAINER}:${PLUGIN_DEST}/"
docker cp "$PLUGIN_SRC" "${CONTAINER}:${PLUGIN_DEST}/"

echo "==> Restart $CONTAINER ..."
docker restart "$CONTAINER"

echo "==> Plugin deployed. Hard-refresh the browser after Document Server is up."
echo "    Expect console: [WebSocket] 使用配置连接: ws://...:19102?type=plugin"
docker logs --tail 20 "$CONTAINER"
