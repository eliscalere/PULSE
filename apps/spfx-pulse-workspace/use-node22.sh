#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
export PATH="$ROOT_DIR/.local/node22/bin:$PATH"
export PATH="$ROOT_DIR/node_modules/.bin:$PATH"

exec "$@"
