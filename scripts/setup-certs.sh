#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CERT_DIR="$ROOT/certs"
mkdir -p "$CERT_DIR"

if command -v mkcert >/dev/null 2>&1; then
  mkcert -install
  mkcert -cert-file "$CERT_DIR/localhost.pem" -key-file "$CERT_DIR/localhost-key.pem" localhost 127.0.0.1 ::1
  echo "Wrote mkcert pair to $CERT_DIR"
  exit 0
fi

echo "mkcert not found. Installing a trusted localhost pair with office-addin-dev-certs."
cd "$ROOT/frontend"
npx --yes office-addin-dev-certs install
DEV_CERTS="${HOME}/.office-addin-dev-certs"
if [[ -f "$DEV_CERTS/localhost.crt" && -f "$DEV_CERTS/localhost.key" ]]; then
  cp "$DEV_CERTS/localhost.crt" "$CERT_DIR/localhost.pem"
  cp "$DEV_CERTS/localhost.key" "$CERT_DIR/localhost-key.pem"
  echo "Copied office-addin-dev-certs pair to $CERT_DIR for uvicorn"
fi
