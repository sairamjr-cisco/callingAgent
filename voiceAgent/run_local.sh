#!/usr/bin/env zsh
set -euo pipefail

ROOT="/Users/sairamjonagadla/Documents/calling-agent/voiceAgent"
ENV_FILE="$ROOT/.env.local"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Missing $ENV_FILE"
  exit 1
fi

cd "$ROOT"

if [[ ! -d ".venv" ]]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q -r requirements.txt
mkdir -p "$ROOT/logs"
mkdir -p "$ROOT/state"
mkdir -p "$ROOT/artifacts"

set -a
source "$ENV_FILE"
set +a

python3 app.py
