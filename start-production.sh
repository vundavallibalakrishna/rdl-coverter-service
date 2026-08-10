#!/usr/bin/env sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$SCRIPT_DIR"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js is not available on PATH." >&2
  exit 1
fi

# scripts/start-production.js loads RDL_ENV_FILE (default: .env.production), applies production defaults
# for values that remain unset, and then starts the service. All application telemetry stays on stdout.
exec node scripts/start-production.js
