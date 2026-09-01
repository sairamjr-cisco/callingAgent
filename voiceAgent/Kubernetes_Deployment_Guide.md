# Kubernetes Production Deployment Guide
**Target Server:** `192.168.7.50`  
**Namespace:** `qway-calling-agent`  

This guide provides step-by-step instructions to securely deploy your clean, lightweight production bundle (`calling-agent-production.zip`) to your Kubernetes cluster in the `qway-calling-agent` namespace.

---

### **Overview of K8s Deployment Architecture**
To make your voice agent fully cloud-native and easily configurable, we separate **application code** (Docker image), **campaign settings** (ConfigMap), and **credentials** (Secrets):
1. **Container Image:** Houses python execution files (`app.py`, `convertExcelToJson.py`) and UI assets.
2. **ConfigMap:** Mounts `campaign.properties` dynamically into the container so you can tweak campaign settings without rebuilding the Docker image!
3. **Persistent Volume Claim (PVC):** Ensures voice recording artifacts and logs persist across Pod restarts.
4. **Secrets:** Securely stores your private Twilio & Gemini credentials.

---

### **Step 1: Upload the Clean Bundle to Server**
Use `scp` to copy the production zip package directly from your local Mac to the deployment gateway:
```bash
scp -P 22 /Users/sairamjonagadla/Documents/calling-agent/calling-agent-production.zip sairam@192.168.7.50:/home/sairam/
```

On your server (`192.168.7.50`), ssh in and unzip the bundle:
```bash
ssh sairam@192.168.7.50
unzip calling-agent-production.zip -d qway-calling-agent/
cd qway-calling-agent/voiceAgent/
```

---

### **Step 2: Define the Production Dockerfile**
Create a highly optimized, clean Dockerfile (`Dockerfile`) in the `voiceAgent/` root:
```dockerfile
FROM python:3.9-slim

# Set system optimization parameters
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=5501

WORKDIR /app

# Install dependencies first (for faster caching layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy essential code and blueprints
COPY . .

# Expose server port
EXPOSE 5501

CMD ["python", "app.py"]
```

---

### **Step 3: Build & Push the Docker Image**
Build the production Docker container and push it to your local registry (e.g. `192.168.7.50:5000`):
```bash
# Build the image
docker build -t 192.168.7.50:5000/qway-calling-agent:latest .

# Push to your server's local container registry
docker push 192.168.7.50:5000/qway-calling-agent:latest
```

---

### **Step 4: Create Kubernetes Secrets for Credentials**
Store your Twilio and Gemini API keys securely inside the `qway-calling-agent` namespace. Run this command on the K8s gateway:
```bash
kubectl create secret generic calling-agent-secrets \
  --namespace=qway-calling-agent \
  --from-literal=TWILIO_ACCOUNT_SID="ACxxxxxxxxxxxxxxxx" \
  --from-literal=TWILIO_AUTH_TOKEN="your_auth_token" \
  --from-literal=TWILIO_PHONE_NUMBER="+12605973908" \
  --from-literal=GEMINI_API_KEY="your_gemini_api_key" \
  --dry-run=client -o yaml | kubectl apply -f -
```

---

### **Step 5: Write the Kubernetes Manifests**
Prepare your deployment manifests inside `voiceAgent/k8s/` directory.

#### **A. ConfigMap (`k8s-configmap.yaml`)**
Holds your editable campaign settings so you can modify directories or rules without rebuilding:
```yaml
apiVersion: v1
kind: ConfigMap
metadata:
  name: calling-agent-properties
  namespace: qway-calling-agent
data:
  campaign.properties: |
    run_excel_converter=true
    call_delay=5
    agent_voice=Polly.Joanna
    payer_profile=aetna_representative
    enable_recording=false
    enable_transcript=true
    storage_mode=server_delete_twilio
    artifact_subfolder=default
    contacts_file=contacts_sheet.xlsx
    script_file=script.json
```

#### **B. Deployment & Service (`k8s-deployment.yaml`)**
Deploys the Flask app container, configures volume mounts for recordings, and exposes ports:
```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: calling-agent
  namespace: qway-calling-agent
  labels:
    app: calling-agent
spec:
  replicas: 1
  selector:
    matchLabels:
      app: calling-agent
  template:
    metadata:
      labels:
        app: calling-agent
    spec:
      containers:
      - name: calling-agent
        image: 192.168.7.50:5000/qway-calling-agent:latest
        imagePullPolicy: Always
        ports:
        - containerPort: 5501
        env:
        # Pull twilio credentials securely from secrets
        - name: TWILIO_ACCOUNT_SID
          valueFrom:
            secretKeyRef:
              name: calling-agent-secrets
              key: TWILIO_ACCOUNT_SID
        - name: TWILIO_AUTH_TOKEN
          valueFrom:
            secretKeyRef:
              name: calling-agent-secrets
              key: TWILIO_AUTH_TOKEN
        - name: TWILIO_PHONE_NUMBER
          valueFrom:
            secretKeyRef:
              name: calling-agent-secrets
              key: TWILIO_PHONE_NUMBER
        - name: GEMINI_API_KEY
          valueFrom:
            secretKeyRef:
              name: calling-agent-secrets
              key: GEMINI_API_KEY
        volumeMounts:
        # Mount the editable properties file dynamically
        - name: properties-volume
          mountPath: /app/campaign.properties
          subPath: campaign.properties
        # Mount PVC to persist call recordings and logs
        - name: storage-volume
          mountPath: /app/artifacts
      volumes:
      - name: properties-volume
        configMap:
          name: calling-agent-properties
      - name: storage-volume
        persistentVolumeClaim:
          claimName: calling-agent-pvc
---
apiVersion: v1
kind: Service
metadata:
  name: calling-agent-service
  namespace: qway-calling-agent
spec:
  selector:
    app: calling-agent
  ports:
  - protocol: TCP
    port: 80
    targetPort: 5501
  type: ClusterIP
```

---

### **Step 6: Apply the Configurations to Namespace**
Deploy everything directly into the cluster inside `qway-calling-agent` namespace:
```bash
# Apply ConfigMap
kubectl apply -f k8s-configmap.yaml

# Apply Deployment & Service
kubectl apply -f k8s-deployment.yaml
```

---

### **Step 7: Verification & Logging**
To verify the deployment is running successfully inside K8s:
```bash
# Check the deployment pods status
kubectl get pods -n qway-calling-agent

# Stream active Flask logs live to verify startup
kubectl logs -f deployment/calling-agent -n qway-calling-agent

# Get cluster endpoint details
kubectl get svc -n qway-calling-agent
```
