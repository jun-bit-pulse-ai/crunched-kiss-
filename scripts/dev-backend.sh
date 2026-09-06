#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CERT="$ROOT/certs/localhost.pem"
KEY="$ROOT/certs/localhost-key.pem"
cd "$ROOT/backend"
# shellcheck disable=SC1091
source .venv/bin/activate

if [[ -f "$CERT" && -f "$KEY" ]]; then
  exec uvicorn app.main:app --reload --host 127.0.0.1 --port 8000 --ssl-certfile "$CERT" --ssl-keyfile "$KEY"
fi

DEV_CERTS="${HOME}/.office-addin-dev-certs"
if [[ -f "$DEV_CERTS/localhost.crt" && -f "$DEV_CERTS/localhost.key" ]]; then
  exec uvicorn app.main:app --reload --host 127.0.0.1 --port 8000 --ssl-certfile "$DEV_CERTS/localhost.crt" --ssl-keyfile "$DEV_CERTS/localhost.key"
fi

echo "No HTTPS certs found. Run ./scripts/setup-certs.sh first." >&2
exit 1
