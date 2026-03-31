#!/usr/bin/env bash

set -euo pipefail

DEPLOY_DIR="${DEPLOY_DIR:-/opt/slack-excel-bot}"
COMPOSE_FILE="${COMPOSE_FILE:-$DEPLOY_DIR/docker-compose.prod.yml}"
ENV_FILE="${ENV_FILE:-$DEPLOY_DIR/.env}"
HOST_DATA_DIR="${HOST_DATA_DIR:-$DEPLOY_DIR/.data}"
RELEASE_FILE="${RELEASE_FILE:-$DEPLOY_DIR/.release.json}"

IMAGE_URI="${IMAGE_URI:?IMAGE_URI is required}"
IMAGE_TAG="${IMAGE_TAG:?IMAGE_TAG is required}"
GITHUB_SHA="${GITHUB_SHA:?GITHUB_SHA is required}"
GITHUB_REF="${GITHUB_REF:?GITHUB_REF is required}"

mkdir -p "$DEPLOY_DIR" "$HOST_DATA_DIR"

if ! command -v docker >/dev/null 2>&1; then
  echo "docker is required on the target host" >&2
  exit 1
fi

if ! docker compose version >/dev/null 2>&1; then
  echo "docker compose plugin is required on the target host" >&2
  exit 1
fi

if [[ ! -f "$COMPOSE_FILE" ]]; then
  echo "Missing compose file at $COMPOSE_FILE" >&2
  exit 1
fi

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Missing env file at $ENV_FILE" >&2
  exit 1
fi

if [[ "$IMAGE_URI" == *"/"* ]]; then
  docker pull "$IMAGE_URI"
fi

IMAGE_URI="$IMAGE_URI" HOST_DATA_DIR="$HOST_DATA_DIR" docker compose -f "$COMPOSE_FILE" up -d --remove-orphans

for legacy in expenses-agent-api expenses-agent-slack-bot; do
  if docker ps -a --format '{{.Names}}' | grep -Fxq "$legacy"; then
    docker stop "$legacy" || true
    docker rm "$legacy" || true
  fi
done

python3 - <<'PY' > "$RELEASE_FILE"
import json
import os
from datetime import datetime, timezone

payload = {
    "image_uri": os.environ["IMAGE_URI"],
    "image_tag": os.environ["IMAGE_TAG"],
    "github_sha": os.environ["GITHUB_SHA"],
    "github_ref": os.environ["GITHUB_REF"],
    "deployed_at": datetime.now(timezone.utc).isoformat(),
}
print(json.dumps(payload, ensure_ascii=False, indent=2))
PY

docker ps --filter name=slack-excel-bot --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}'
