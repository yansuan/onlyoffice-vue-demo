#!/usr/bin/env bash
# Deploy jsapi-executor plugin into Document Server and restart.
# Same effect as docs/docker.txt plugin copy steps.
#
# Usage:
#   ./scripts/oo.sh           # copy, restart, follow logs
#   ./scripts/oo.sh --no-logs # copy, restart, exit
set -euo pipefail
cd "$(dirname "$0")/.."

CONTAINER="${CONTAINER:-documentserver}"
PLUGIN_SRC="plugins/jsapi-executor"
PLUGIN_DEST="/var/www/onlyoffice/documentserver/sdkjs-plugins"
FOLLOW_LOGS=1

for arg in "$@"; do
  case "$arg" in
    --no-logs|-n) FOLLOW_LOGS=0 ;;
    -h|--help)
      echo "Usage: $0 [--no-logs]"
      echo "  CONTAINER=documentserver  Document Server container name"
      exit 0
      ;;
    *)
      echo "Unknown option: $arg" >&2
      exit 1
      ;;
  esac
done

if ! docker inspect "$CONTAINER" >/dev/null 2>&1; then
  echo "Container '$CONTAINER' not found. Start Document Server first (see docs/docker.txt)." >&2
  exit 1
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

if [[ "$FOLLOW_LOGS" -eq 1 ]]; then
  docker logs "$CONTAINER" -f
fi
