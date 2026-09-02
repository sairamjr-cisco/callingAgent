#!/usr/bin/env zsh
set -euo pipefail

ROOT="/Users/sairamjonagadla/Documents/calling-agent"
SRC_DIR="/Users/sairamjonagadla/Documents/calling-agent/voiceAgent"
TEMP_DIR="$ROOT/temp_deploy"
ZIP_OUT="$ROOT/calling-agent-production.zip"

echo "=== Preparing clean production bundle ==="

# Clean old zip if exists
if [[ -f "$ZIP_OUT" ]]; then
  rm -f "$ZIP_OUT"
fi

# Create clean temp directory
mkdir -p "$TEMP_DIR/voiceAgent"
mkdir -p "$TEMP_DIR/voiceAgent/prompts"
mkdir -p "$TEMP_DIR/voiceAgent/input"
mkdir -p "$TEMP_DIR/voiceAgent/static"

# Copy core python & configuration scripts
cp "$SRC_DIR/app.py" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/convertExcelToJson.py" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/campaign.properties" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/payer_profiles.json" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/requirements.txt" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/index.html" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/results.html" "$TEMP_DIR/voiceAgent/"
cp "$SRC_DIR/run_local.sh" "$TEMP_DIR/voiceAgent/"

# Copy prompt files dynamically
if [ -d "$SRC_DIR/prompts" ]; then
  cp -R "$SRC_DIR/prompts/" "$TEMP_DIR/voiceAgent/prompts/"
fi

# Copy input Excel files dynamically
if [ -d "$SRC_DIR/input" ]; then
  cp -R "$SRC_DIR/input/" "$TEMP_DIR/voiceAgent/input/"
fi

# Copy UI elements (static assets)
cp -R "$SRC_DIR/static/" "$TEMP_DIR/voiceAgent/static/"

# Create empty target directories for server runtime logs and logs mapping
mkdir -p "$TEMP_DIR/voiceAgent/logs"
mkdir -p "$TEMP_DIR/voiceAgent/state"
mkdir -p "$TEMP_DIR/voiceAgent/artifacts"

# Build Zip Archive
cd "$TEMP_DIR"
zip -q -r "$ZIP_OUT" voiceAgent/

# Cleanup temporary staging folder
rm -rf "$TEMP_DIR"

echo "=== Success! Production package ready at: ==="
echo "📂 $ZIP_OUT"
