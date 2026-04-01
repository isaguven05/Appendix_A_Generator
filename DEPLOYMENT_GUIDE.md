# Appendix Automation Tool — Local Server Deployment Guide

**For Metricell IT / the person setting this up.**
This guide covers everything needed to get the tool running on a local Windows or Linux server so the whole team can access it from any browser on the internal network.

---

## What this tool is

A Python web application (Flask) that generates PowerPoint appendix files for mobile network projects. Users visit a web page in their browser, upload their files, and download ready-made PPTX decks. No PowerPoint install is needed on the server.

---

## Prerequisites

| Requirement | Minimum version | Notes |
|---|---|---|
| Python | 3.9 or later | Download from python.org |
| pip | bundled with Python | Used to install dependencies |
| Git | any recent version | To pull the code from GitHub |
| Network access | — | Server must be reachable on your LAN |

---

## Step 1 — Install Python

### Windows
1. Go to https://www.python.org/downloads/
2. Download the latest Python 3.x installer
3. **Important:** tick **"Add Python to PATH"** during installation
4. Open Command Prompt and verify:
   ```
   python --version
   ```
   You should see something like `Python 3.11.9`

### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv git -y
python3 --version
```

---

## Step 2 — Download the application code

On the server, open a terminal (or Command Prompt on Windows) and run:

```bash
git clone https://github.com/isaguven05/Appendix_A_Generator.git
cd Appendix_A_Generator
```

This downloads all the code into a folder called `Appendix_A_Generator`.

> **No Git on the server?**
> You can also download a ZIP from GitHub: go to the repo page → green **Code** button → **Download ZIP**, then extract it to a folder on the server.

---

## Step 3 — Create a virtual environment

A virtual environment keeps the app's dependencies isolated from the rest of the system.

### Windows
```cmd
python -m venv venv
venv\Scripts\activate
```

### Linux / macOS
```bash
python3 -m venv venv
source venv/bin/activate
```

You should now see `(venv)` at the start of your prompt.

---

## Step 4 — Install dependencies

With the virtual environment active:

```bash
pip install -r requirements.txt
```

This installs Flask, python-pptx, Pillow, lxml, and the production server (Gunicorn on Linux, Waitress on Windows). It may take a minute or two.

---

## Step 5 — Test it locally first

Run the app in development mode just to confirm everything works:

### Windows
```cmd
python app.py
```

### Linux
```bash
python3 app.py
```

Open a browser on the **same machine** and go to:
```
http://127.0.0.1:5000
```

You should see the Metricell Appendix Automation Tool home page. If it loads, stop the server (`Ctrl+C`) and move to Step 6.

---

## Step 6 — Run with a production server

The built-in Flask dev server is not suitable for real use. Use one of the following instead.

### Option A — Windows (Waitress) ✅ Recommended for Windows servers

```cmd
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

### Option B — Linux (Gunicorn) ✅ Recommended for Linux servers

```bash
gunicorn --workers 2 --bind 0.0.0.0:5000 app:app
```

`--workers 2` means two processes can handle requests simultaneously. You can increase this if many people use it at once.

The `--bind 0.0.0.0:5000` part makes the app accessible from other computers on the network (not just localhost).

---

## Step 7 — Access from other computers

Find the server's local IP address:

### Windows
```cmd
ipconfig
```
Look for **IPv4 Address** under your network adapter — e.g. `192.168.1.50`

### Linux
```bash
hostname -I
```

Now any computer on the same network can open a browser and go to:
```
http://192.168.1.50:5000
```

> **Tip:** Ask IT to assign the server a fixed/static IP so the address never changes.

---

## Step 8 — Keep it running permanently (optional but recommended)

You don't want to have to manually start the app every time the server reboots.

### Windows — run as a background service using NSSM

1. Download NSSM from https://nssm.cc/download
2. Open an **Administrator** Command Prompt and run:
   ```cmd
   nssm install AppendixTool
   ```
3. In the NSSM window that opens:
   - **Path:** `C:\path\to\Appendix_A_Generator\venv\Scripts\waitress-serve.exe`
   - **Arguments:** `--host=0.0.0.0 --port=5000 app:app`
   - **Startup directory:** `C:\path\to\Appendix_A_Generator`
4. Click **Install service**, then start it:
   ```cmd
   nssm start AppendixTool
   ```

The tool will now start automatically when the server boots.

### Linux — run as a systemd service

Create a service file:

```bash
sudo nano /etc/systemd/system/appendix-tool.service
```

Paste the following (replace `/home/username/Appendix_A_Generator` with your actual path):

```ini
[Unit]
Description=Metricell Appendix Automation Tool
After=network.target

[Service]
User=www-data
WorkingDirectory=/home/username/Appendix_A_Generator
ExecStart=/home/username/Appendix_A_Generator/venv/bin/gunicorn --workers 2 --bind 0.0.0.0:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start it:
```bash
sudo systemctl daemon-reload
sudo systemctl enable appendix-tool
sudo systemctl start appendix-tool
sudo systemctl status appendix-tool   # should show "active (running)"
```

---

## Step 9 — (Optional) Give it a nice URL instead of an IP

Ask your IT team to add a local DNS entry so users can type something like:

```
http://appendix-tool/
```

instead of the IP address. This is a simple internal DNS record pointing the hostname to the server's IP.

Alternatively, if you have an internal Nginx or IIS reverse proxy, point it at `127.0.0.1:5000`.

---

## Step 10 — Firewall

Make sure the server's firewall allows inbound connections on **port 5000** from the internal network.

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

When a new version is pushed to GitHub, updating is a single command:

```bash
cd Appendix_A_Generator
git pull
```

Then restart the service:

```bash
# Linux
sudo systemctl restart appendix-tool

# Windows
nssm restart AppendixTool
```

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| Browser says "This site can't be reached" | Firewall blocking port 5000 | See Step 10 |
| `ModuleNotFoundError` when starting | Dependencies not installed | Run `pip install -r requirements.txt` with venv active |
| App starts but PPTX generation fails | Template files missing | Make sure `templateB.pptx` and `template_A.pptx` are in the project folder |
| Port 5000 already in use | Another process is using it | Change `--port=5000` to `--port=5001` (or any free port) |
| `Permission denied` on Linux | Wrong file ownership | Run `sudo chown -R www-data:www-data /path/to/Appendix_A_Generator` |

---

## Summary — quick reference

```
# One-time setup
git clone https://github.com/isaguven05/Appendix_A_Generator.git
cd Appendix_A_Generator
python3 -m venv venv && source venv/bin/activate   # Linux
pip install -r requirements.txt

# Start (Linux)
gunicorn --workers 2 --bind 0.0.0.0:5000 app:app

# Start (Windows)
waitress-serve --host=0.0.0.0 --port=5000 app:app

# Access from any machine on the network
http://<server-ip>:5000
```

---

*Prepared for Metricell internal use. Contact the original developer for application-level support.*
