#!/usr/bin/env zsh
set -euo pipefail

# ==========================================
# CONFIGURATION SETTINGS
# ==========================================
SERVER_IP="192.168.7.50"
SERVER_USER="sairam"
PORT="22"
KEY_FILE="/Users/sairamjonagadla/sairam.key"
NAMESPACE="qway-calling-agent"
DEPLOYMENT="calling-agent"
REMOTE_DIR="/home/$SERVER_USER/qway-calling-agent"

ROOT="/Users/sairamjonagadla/Documents/calling-agent"
ZIP_FILE="$ROOT/calling-agent-production.zip"
PACKAGER="$ROOT/package_production.sh"

echo "=================================================="
echo "🚀 AUTOMATED DEPLOYMENT FOR QWAY CALLING AGENT"
echo "=================================================="
echo "Target Host: $SERVER_USER@$SERVER_IP (Namespace: $NAMESPACE)"
echo "SSH Identity Key: $KEY_FILE"

# Verify that the private key file exists
if [[ ! -f "$KEY_FILE" ]]; then
  echo "❌ Error: SSH identity key file not found at: $KEY_FILE"
  exit 1
fi

# 1. Compile the latest production package locally
if [[ -f "$PACKAGER" ]]; then
  echo "📦 Step 1: Compiling clean production package..."
  zsh "$PACKAGER"
else
  echo "❌ Error: Packaging script $PACKAGER not found!"
  exit 1
fi

if [[ ! -f "$ZIP_FILE" ]]; then
  echo "❌ Error: Production bundle $ZIP_FILE was not compiled!"
  exit 1
fi

# 2. Upload the production zip package to the server gateway
echo "📤 Step 2: Uploading production bundle to server..."
echo "Running: scp -i \"$KEY_FILE\" -P $PORT \"$ZIP_FILE\" \"$SERVER_USER@$SERVER_IP:/home/$SERVER_USER/\""
scp -i "$KEY_FILE" -P "$PORT" "$ZIP_FILE" "$SERVER_USER@$SERVER_IP:/home/$SERVER_USER/"

# 3. Create the ConfigMap locally from current campaign.properties
echo "📝 Step 3: Preparing dynamic Kubernetes ConfigMap..."
PROPERTIES_FILE="/Users/sairamjonagadla/Documents/voiceAgent_working_31_Aug/campaign.properties"
CONFIGMAP_YAML="/Users/sairamjonagadla/Documents/voiceAgent_working_31_Aug/k8s-configmap.yaml"

if [[ -f "$PROPERTIES_FILE" ]]; then
  cat <<EOF > "$CONFIGMAP_YAML"
apiVersion: v1
kind: ConfigMap
metadata:
  name: calling-agent-properties
  namespace: $NAMESPACE
data:
  campaign.properties: |
$(sed 's/^/    /' "$PROPERTIES_FILE")
EOF
  echo "✅ Staged K8s ConfigMap manifest at: /Users/sairamjonagadla/Documents/voiceAgent_working_31_Aug/k8s-configmap.yaml"
else
  echo "⚠️ Warning: campaign.properties not found! Skipping ConfigMap generation."
fi

# Upload the ConfigMap manifest to the server
if [[ -f "$CONFIGMAP_YAML" ]]; then
  scp -i "$KEY_FILE" -P "$PORT" "$CONFIGMAP_YAML" "$SERVER_USER@$SERVER_IP:/home/$SERVER_USER/"
fi

# 4. Trigger remote SSH commands to unpack, build Docker image, and rollout update
echo "⚙️ Step 4: Executing remote deployment commands over SSH..."
ssh -i "$KEY_FILE" -p "$PORT" "$SERVER_USER@$SERVER_IP" <<EOF
  set -euo pipefail
  
  echo "   -> Unpacking zip package..."
  mkdir -p "$REMOTE_DIR"
  unzip -q -o "/home/$SERVER_USER/calling-agent-production.zip" -d "$REMOTE_DIR/"
  cd "$REMOTE_DIR/voiceAgent/"

  # If a Dockerfile doesn't exist on the server, create it dynamically
  if [ ! -f Dockerfile ]; then
    echo "   -> Creating production Dockerfile..."
    cat <<DOCKER > Dockerfile
FROM python:3.9-slim
ENV PYTHONDONTWRITEBYTECODE=1 \\
    PYTHONUNBUFFERED=1 \\
    PORT=5501
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 5501
CMD ["python", "app.py"]
DOCKER
  fi

  echo "   -> Building production Docker container..."
  docker build -t qwayhealthcare/qway-calling-agent:latest .

  echo "   -> Pushing container image to Docker Hub registry..."
  docker push qwayhealthcare/qway-calling-agent:latest

  echo "   -> Applying ConfigMap configurations..."
  if [ -f "/home/$SERVER_USER/k8s-configmap.yaml" ]; then
    kubectl apply -f "/home/$SERVER_USER/k8s-configmap.yaml"
    rm -f "/home/$SERVER_USER/k8s-configmap.yaml"
  fi

  echo "   -> Triggering rolling rollout update in K8s (Zero-Downtime)..."
  kubectl rollout restart deployment/$DEPLOYMENT -n $NAMESPACE

  echo "   -> Cleaning up staging artifacts..."
  rm -f "/home/$SERVER_USER/calling-agent-production.zip"
EOF

# 5. Output cluster active status details
echo "=================================================="
echo "🎉 DEPLOYMENT COMMANDS SENT SUCCESSFULLY!"
echo "=================================================="
echo "Verify status on server using: kubectl get pods -n $NAMESPACE"
