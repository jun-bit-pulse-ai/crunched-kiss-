#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT/backend"
# shellcheck disable=SC1091
source .venv/bin/activate

# HTTP only. Excel talks to https://localhost:3000/api/* and webpack proxies here.
exec uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
