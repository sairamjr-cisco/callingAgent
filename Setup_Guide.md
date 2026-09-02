# Setup Guide: Qway Calling Agent (Local & Development Setup)

This guide walks you through setting up and running the Qway Calling Agent on a brand-new laptop for local development, testing, and production deployment.

---

## 1. Prerequisites

Before starting, ensure the following software is installed on your laptop:

- **Python 3.9+** (Check with `python3 --version`)
- **Git** (Check with `git --version`)
- **A Tunneling Tool** (e.g., [Cloudflare Tunnel / Cloudflared](https://developers.cloudflare.com/pages/how-to/preview-with-cloudflare-tunnel/) or [Ngrok](https://ngrok.com/)) to receive Twilio Webhooks locally.
- **Docker** & **Kubectl** (Optional, only required if building production images or managing Kubernetes).

---

## 2. Clone and Structure Overview

Clone your repository to the local machine:

```bash
git clone https://github.com/sairamjr-cisco/callingAgent.git
cd callingAgent
```

### Directory Structure of Interest:
- `voiceAgent/` - Central Python/Flask backend and frontend assets.
  - `app.py` - Flask web application, speech hook, API definitions.
  - `requirements.txt` - Python package dependencies.
  - `payer_profiles.json` - Payer telephone and workbook mapping configuration.
  - `prompts/` - JSON configurations for payer-specific welcome-tree rules (e.g., UHC, Aetna).
  - `input/` - Contains the runtime runtime `customer_data.json` and templates.
  - `static/` - UI stylesheets and script behaviors (`app.js`, `results.js`).
  - `templates/` - HTML files.
- `package_production.sh` - Packages the project into a deployment bundle zip.
- `deploy_to_server.sh` - Standard deployment executor script (SSH & Kubernetes).

---

## 3. Configuration Setup (`.env.local`)

Inside the `voiceAgent/` directory, create a local environment file named `.env.local` to manage secrets and system paths.

Create `/Users/sairamjonagadla/Documents/calling-agent/voiceAgent/.env.local` with the following variables:

```ini
PORT=5501
PUBLIC_BASE_URL=https://your-tunnel-subdomain.trycloudflare.com

# Twilio Credentials
TWILIO_ACCOUNT_SID=ACXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
TWILIO_AUTH_TOKEN=your_twilio_auth_token_here
TWILIO_PHONE_NUMBER=+1234567890
TWILIO_WEBHOOK_SIGNATURE_VALIDATION=true

# AI / Extraction Safety Net
GEMINI_API_KEY=your_gemini_api_key_here

# Default tuning
DEFAULT_VOICE=Polly.Joanna
CUSTOMER_CARE_SPEECH_TIMEOUT_SECONDS=2

# Directory & Log Paths (Adjust absolute path to match your new laptop workspace)
APP_LOG_FILE=/path/to/calling-agent/voiceAgent/logs/calling-agent.log
CALL_STATE_DIR=/path/to/calling-agent/voiceAgent/state
CALL_ARTIFACTS_DIR=/path/to/calling-agent/voiceAgent/artifacts

# Reference File Paths
CUSTOMER_DATA_FILE=/path/to/calling-agent/voiceAgent/input/customer_data.json
PAYER_PROFILE_FILE=/path/to/calling-agent/voiceAgent/payer_profiles.json
```

---

## 4. Local Run & Automatic Environment Provisioning

The project includes an automatic initialization script `run_local.sh` that checks configurations, provisions a virtual environment, installs requirements, handles storage folder setup, and launches the application.

To start the app:

```bash
cd voiceAgent
chmod +x run_local.sh
./run_local.sh
```

### Manual Virtual Environment (Backup Walkthrough):
If you prefer running commands manually instead of using `run_local.sh`:

```bash
cd voiceAgent
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

mkdir -p logs state artifacts

# Run application using local env variables
export $(cat .env.local | xargs)
python3 app.py
```

---

## 5. Webhook Tunneling (Receiving Twilio callbacks)

Since Twilio interacts with the calling agent via HTTP webhooks, your local Flask server must be exposed to the internet.

### Using Cloudflare Tunnel (Free & No Signup Needed)
1. Install cloudflared: `brew install cloudflare/cloudflare/cloudflared` (macOS)
2. Expose Flask's port (default: 5501):
   ```bash
   cloudflared tunnel --url http://localhost:5501
   ```
3. Copy the generated `.trycloudflare.com` URL (e.g. `https://indoor-situations-download-one.trycloudflare.com`) and paste it as your `PUBLIC_BASE_URL` in `voiceAgent/.env.local`.
4. Restart your Flask server to register the new webhook endpoints.

---

## 6. Accessing the Dashboard UI

Once the application is running:

1. **Campaign Launcher Panel:** Open `http://localhost:5501/` in your browser.
   - Upload customer Excel spreadsheets per payer (e.g. Aetna, UHC). It converts them to JSON immediately and handles de-duplication automatically.
   - Start batch calling campaigns.

2. **Audit & Results Dashboard:** Access `http://localhost:5501/results`.
   - View ongoing and historical call sessions.
   - Play audio recordings and inspect clean interaction transcripts.
   - View local regex extracted JSON claim audits.
   - Select multiple campaigns and export them to a consolidated Excel spreadsheet in one click.

---

## 7. Formatting & Guidelines
- Make sure standard directories (`logs/`, `state/`, `artifacts/`, `input/`) are present in `.gitignore` so they are not committed to production.
- Excel spreadsheet templates should follow the headers present in `voiceAgent/input/sample_file.xlsx` to ensure successful parsing.
