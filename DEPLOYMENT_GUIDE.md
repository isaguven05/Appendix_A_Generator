# Deployment Guide — Metricell Appendix Automation Tool

**Audience:** Metricell IT or whoever is setting this up on the internal server.

This guide gets the tool running on a local Windows or Linux server so the whole team can access it from any browser on the internal network. Once it's running you don't need to touch it again — it just works in the background.

---

## Overview

- The app is a Python web application (Flask)
- It runs on one server; everyone else just uses their browser — nothing to install on user machines
- Estimated setup time: **20–30 minutes**

---

## Prerequisites

| What | Version | Where to get it |
|---|---|---|
| Python | 3.9 or later | https://www.python.org/downloads/ |
| Git | Any recent version | https://git-scm.com/download/win (Windows) |

---

## Step 1 — Install Python

### Windows
1. Go to https://www.python.org/downloads/ and download the latest Python 3.x installer
2. Run the installer — **tick "Add Python to PATH"** before clicking Install
3. Open Command Prompt and check it worked:
   ```
   python --version
   ```
   You should see something like `Python 3.11.9`

### Linux (Ubuntu/Debian)
```bash
sudo apt update && sudo apt install python3 python3-pip python3-venv git -y
python3 --version
```

---

## Step 2 — Install Git (Windows only)

If `git` is not recognised as a command:

1. Go to https://git-scm.com/download/win
2. Download and run the installer — click through with all default settings
3. Close and reopen Command Prompt, then check:
   ```
   git --version
   ```

> **No internet on the server?** Skip Git. Go to https://github.com/isaguven05/Appendix_A_Generator on another machine, click the green **Code** button → **Download ZIP**, copy the ZIP to the server and extract it. Then skip to Step 4.

---

## Step 3 — Download the application code

Open Command Prompt (Windows) or Terminal (Linux) and run:

```bash
git clone https://github.com/isaguven05/Appendix_A_Generator.git
cd Appendix_A_Generator
```

Choose a sensible permanent location — for example `C:\AppendixTool\` on Windows or `/opt/appendix-tool/` on Linux. Move the folder there now if needed.

---

## Step 4 — Create a virtual environment

This keeps the app's dependencies isolated and avoids conflicts with anything else on the server.

### Windows
```cmd
python -m venv venv
venv\Scripts\activate
```

### Linux
```bash
python3 -m venv venv
source venv/bin/activate
```

You should see `(venv)` appear at the start of your prompt.

---

## Step 5 — Install dependencies

With the virtual environment active:

```bash
pip install -r requirements.txt
```

This installs everything the app needs (Flask, python-pptx, Pillow, lxml, Waitress/Gunicorn). Takes 1–2 minutes.

---

## Step 6 — Test it

Run the app in test mode to confirm everything is working:

### Windows
```cmd
python app.py
```

### Linux
```bash
python3 app.py
```

Open a browser **on the same machine** and go to:
```
http://127.0.0.1:5000
```

You should see the Metricell Appendix Automation Tool home page. If it loads correctly, press `Ctrl+C` to stop it and move on.

---

## Step 7 — Run with a production server

The built-in test server is not suitable for real use. Use the following instead.

### Windows — Waitress
```cmd
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

### Linux — Gunicorn
```bash
gunicorn --workers 2 --bind 0.0.0.0:5000 app:app
```

The `0.0.0.0` part makes the app reachable from other computers on the network.

---

## Step 8 — Find the server's IP address

### Windows
```cmd
ipconfig
```
Look for **IPv4 Address** under your network adapter — e.g. `192.168.1.50`

### Linux
```bash
hostname -I
```

Any computer on the same network can now open a browser and go to:
```
http://192.168.1.50:5000
```

> **Recommended:** Ask IT to assign the server a static IP so this address never changes. You can also set up an internal DNS entry (e.g. `http://appendix-tool/`) so users don't need to remember an IP.

---

## Step 9 — Keep it running permanently

Set the app up as a background service so it starts automatically when the server reboots.

### Windows — using NSSM (Non-Sucking Service Manager)

1. Download NSSM from https://nssm.cc/download and extract it somewhere (e.g. `C:\Tools\nssm.exe`)
2. Open an **Administrator** Command Prompt and run:
   ```cmd
   C:\Tools\nssm.exe install AppendixTool
   ```
3. Fill in the NSSM window:
   - **Path:** full path to `waitress-serve.exe` inside your venv, e.g.
     `C:\AppendixTool\Appendix_A_Generator\venv\Scripts\waitress-serve.exe`
   - **Arguments:** `--host=0.0.0.0 --port=5000 app:app`
   - **Startup directory:** full path to the project folder, e.g.
     `C:\AppendixTool\Appendix_A_Generator`
4. Click **Install service**
5. Start it:
   ```cmd
   C:\Tools\nssm.exe start AppendixTool
   ```

The tool now starts automatically on boot.

### Linux — using systemd

Create a service file:
```bash
sudo nano /etc/systemd/system/appendix-tool.service
```

Paste the following — **replace the paths with your actual install location:**
```ini
[Unit]
Description=Metricell Appendix Automation Tool
After=network.target

[Service]
User=www-data
WorkingDirectory=/opt/appendix-tool/Appendix_A_Generator
ExecStart=/opt/appendix-tool/Appendix_A_Generator/venv/bin/gunicorn --workers 2 --bind 0.0.0.0:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start it:
```bash
sudo systemctl daemon-reload
sudo systemctl enable appendix-tool
sudo systemctl start appendix-tool
sudo systemctl status appendix-tool
```

The last command should show `active (running)`.

---

## Step 10 — Open the firewall

The server needs to allow inbound connections on port 5000.

### Windows Firewall
```cmd
netsh advfirewall firewall add rule name="Appendix Tool" dir=in action=allow protocol=TCP localport=5000
```

### Linux (ufw)
```bash
sudo ufw allow 5000/tcp
```

---

## Updating the app

When a new version is available, updating takes about 30 seconds:

```bash
cd /path/to/Appendix_A_Generator
git pull
pip install -r requirements.txt   # only needed if dependencies changed
```

Then restart the service:
```bash
# Linux
sudo systemctl restart appendix-tool

# Windows
C:\Tools\nssm.exe restart AppendixTool
```

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| "This site can't be reached" from another PC | Firewall blocking port 5000 | See Step 10 |
| `ModuleNotFoundError` on startup | Dependencies not installed | Run `pip install -r requirements.txt` with venv active |
| PPTX generation fails | Template files missing | Check `templateB.pptx` and `template_A_*.pptx` are in the project folder |
| Port 5000 already in use | Another process is on that port | Change `--port=5000` to `--port=5001` throughout |
| `Permission denied` (Linux) | Wrong file ownership | `sudo chown -R www-data:www-data /path/to/Appendix_A_Generator` |
| `git` not recognised | Git not installed | See Step 2 |
| `python` not recognised | Python not installed or not in PATH | See Step 1 — make sure "Add to PATH" was ticked |

---

## Quick reference

```bash
# ── First-time setup ──────────────────────────────────────────
git clone https://github.com/isaguven05/Appendix_A_Generator.git
cd Appendix_A_Generator
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

# ── Start (Linux) ─────────────────────────────────────────────
gunicorn --workers 2 --bind 0.0.0.0:5000 app:app

# ── Start (Windows) ───────────────────────────────────────────
waitress-serve --host=0.0.0.0 --port=5000 app:app

# ── Access from network ───────────────────────────────────────
http://<server-ip>:5000

# ── Update ────────────────────────────────────────────────────
git pull && sudo systemctl restart appendix-tool
```

---

*For application-level issues contact the original developer. For infrastructure issues contact Metricell IT.*
