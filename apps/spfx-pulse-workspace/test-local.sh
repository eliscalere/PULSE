#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
export PATH="$ROOT_DIR/.local/node22/bin:$PATH"

cd "$ROOT_DIR"
npm run env:check
npm run typecheck
npm run build
