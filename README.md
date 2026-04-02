# Metricell Appendix Automation Tool

A web application that automatically generates PowerPoint appendix decks for mobile network deployment projects. Users fill in a short form, upload their files, and download a ready-made `.pptx` in seconds — replacing a process that previously took around 2.5 hours manually.

---

## What it generates

| Tool | Output |
|---|---|
| **Appendix A** | RF Schematics & GA Drawings deck — one file per provider |
| **Appendix B** | RF Predictions deck — one file per provider, extracted from an input PPTX |

---

## How it works

1. A user opens the tool in their browser (no install required on their machine)
2. They fill in metadata (dates, author, station address etc.) and upload their files
3. The tool generates a fully populated `.pptx` per provider and downloads it automatically

---

## Deployment

This tool is designed to run on a **Metricell internal server**. Once running, anyone on the network can access it from any browser — no setup needed on their end.

**See [`DEPLOYMENT_GUIDE.md`](./DEPLOYMENT_GUIDE.md) for full step-by-step instructions.**

Quick summary:

```bash
# 1. Download the code
git clone https://github.com/isaguven05/Appendix_A_Generator.git
cd Appendix_A_Generator

# 2. Create a virtual environment and install dependencies
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

# 3. Start the server
# Linux:
gunicorn --workers 2 --bind 0.0.0.0:5000 app:app
# Windows:
waitress-serve --host=0.0.0.0 --port=5000 app:app

# 4. Open in any browser on the network
http://<server-ip>:5000
```

---

## Project structure

```
Appendix_A_Generator/
├── app.py                  ← Flask routes and request handling
├── generator.py            ← Appendix A PPTX generation logic
├── generator_b.py          ← Appendix B PPTX generation logic
├── engine_b.py             ← Appendix B core engine (slide building, extraction)
├── requirements.txt        ← Python dependencies
├── templateB.pptx          ← Appendix B slide template (do not rename)
├── template_A_EE.pptx      ← Appendix A template — EE
├── template_A_THREE.pptx   ← Appendix A template — Three
├── template_A_VODA.pptx    ← Appendix A template — Vodafone
├── template_A_VMO2.pptx    ← Appendix A template — VMO2
├── templates/              ← HTML pages (home, Appendix A form, Appendix B form)
├── static/                 ← Static assets (CSS, logo)
├── uploads/                ← Temporary upload storage (auto-created)
└── outputs/                ← Generated files (auto-created)
```

---

## Updating the app

When a new version is pushed to GitHub, updating takes one command on the server:

```bash
cd Appendix_A_Generator
git pull
sudo systemctl restart appendix-tool   # Linux
nssm restart AppendixTool              # Windows
```

---

## Tech stack

- **Python 3.9+**
- **Flask** — web framework
- **python-pptx** — PowerPoint generation
- **lxml** — XML manipulation for advanced PPTX operations
- **Pillow** — image handling
- **Gunicorn** (Linux) / **Waitress** (Windows) — production WSGI server

---

## Support

For issues with the application, contact the original developer.
For server or infrastructure issues, contact Metricell IT.
