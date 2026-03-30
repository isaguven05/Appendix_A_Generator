# Appendix A – PPTX Generator

A web application that automatically builds the **Appendix A – RF Schematics & GA Drawings** PowerPoint from a template. Fill in a short form, upload your files, and download a fully populated `.pptx` in seconds.

---

## Option A — Share a Public Link (No Setup for Anyone)

Deploy to **Railway** — free, takes ~5 minutes, and gives you a permanent link like:
```
https://appendix-a.up.railway.app
```
Anyone can open it in a browser with no installation required.

### Step 1 — Put the project on GitHub
1. Go to [github.com](https://github.com) and create a free account if you don't have one
2. Click the **+** icon → **New repository** → name it `appendix-a-automation` → click **Create repository**
3. Download [GitHub Desktop](https://desktop.github.com/) (easiest way to upload files)
4. Open GitHub Desktop → **Add an Existing Repository** → select the `APPENDIX_A_AUTOMATION` folder
5. Click **Publish repository** → make sure **Keep this code private** is ticked → click **Publish**

> The `uploads/` and `outputs/` folders are automatically excluded via `.gitignore`.

### Step 2 — Deploy on Railway
1. Go to [railway.app](https://railway.app) and sign up with your GitHub account
2. Click **New Project** → **Deploy from GitHub repo**
3. Select your `appendix-a-automation` repository
4. Railway detects Python automatically and deploys — wait ~2 minutes
5. Click **Settings** → **Networking** → **Generate Domain**
6. Copy your public URL and share it with anyone

### Step 3 — Updating the App Later
Whenever you change any code:
1. Open GitHub Desktop
2. Write a short message (e.g. "fix building title") and click **Commit**
3. Click **Push origin**
4. Railway automatically redeploys within ~1 minute — the link stays the same

> **Cost:** Free for hobby use (500 hours/month). No credit card required.

---

## Option B — Run Locally (Just You)

Use this if you only need the app on your own computer.

---

## What You Need Before Starting

| Item | Details |
|---|---|
| **Python 3.9 or newer** | [python.org/downloads](https://www.python.org/downloads/) |
| **`template_A.pptx`** | Must be in the same folder as `app.py` |
| **Input P01 `.pptx`** | The existing P01 version of the pack |
| **3D image** | File named exactly `3D_STATIONCODE_STATIONNAME` (e.g. `3D_DO93_TEMPLE.png`) |
| **Design plan images** | Named containing `"Design plan"` + a number (e.g. `EE_TEMPLE - Design plan1.jpeg`) |
| **Building (GA) images** | Named as `ANYTHING - Building X - LOCATION` (e.g. `EE_TEMPLE - Building 1 - Midway.jpeg`) |

> **Image naming matters.** The app reads the slide title directly from the filename, so keep names accurate.

---

## Step 1 (Local) — Install Python

### Mac
1. Open **Terminal** (press `Cmd + Space`, type `Terminal`, press Enter)
2. Check if Python is already installed:
   ```
   python3 --version
   ```
   If you see `Python 3.x.x` you're good. If not, download it from [python.org](https://www.python.org/downloads/).

### Windows
1. Open **Command Prompt** (press `Win + R`, type `cmd`, press Enter)
2. Check if Python is already installed:
   ```
   python --version
   ```
   If you see `Python 3.x.x` you're good. If not, download it from [python.org](https://www.python.org/downloads/).
   > During installation, **tick the box that says "Add Python to PATH"** — this is important!

---

## Step 2 (Local) — Install the Required Libraries

Open your terminal / command prompt and navigate to the project folder.

### Mac
```bash
cd ~/Downloads/APPENDIX_A_AUTOMATION
pip3 install flask python-pptx Pillow werkzeug
```

### Windows
```cmd
cd %USERPROFILE%\Downloads\APPENDIX_A_AUTOMATION
pip install flask python-pptx Pillow werkzeug
```

> You only need to do this **once**. After the first install, skip straight to Step 3 next time.

---

## Step 3 (Local) — Start the App

### Mac
```bash
cd ~/Downloads/APPENDIX_A_AUTOMATION
python3 app.py
```

### Windows
```cmd
cd %USERPROFILE%\Downloads\APPENDIX_A_AUTOMATION
python app.py
```

You should see something like:
```
* Running on http://127.0.0.1:5050
```

---

## Step 4 (Local) — Open the App in Your Browser

Open any web browser (Chrome, Edge, Safari) and go to:

```
http://localhost:5050
```

---

## Step 5 (Local & Hosted) — Fill in the Form & Upload Files

| Field | What to enter |
|---|---|
| **P01 Date** | Date of the original submission (e.g. `24/01/2025`) |
| **P02 Date** | Date of the updated design (e.g. `25/03/2025`) |
| **Author** | Full name of the person who drew it |
| **Checked** | Full name of the checker |
| **Approved** | Full name of the approver |
| **Station Address** | Full station address |
| **3D Image** | The single 3D render file (`3D_STATIONCODE_NAME.png`) |
| **Input PPTX (P01)** | The existing P01 pack (used to count GA slides) |
| **Design & Building Images** | Select **all** images at once — design plans and building images together |

> **Tip (Mac):** Hold `Cmd` to select multiple files.
> **Tip (Windows):** Hold `Ctrl` to select multiple files.

Click **Generate PPTX** and wait a few seconds. The file will download automatically.

---

## File Naming Rules (Important)

The app sorts and categorises images **purely by their filename**, so follow these patterns:

### Design Plan Images
Must contain `"Design plan"` followed by a number:
```
EE_TEMPLE - Design plan1.jpeg
EE_TEMPLE - Design plan2.jpeg
EE_TEMPLE - Design plan3.jpeg
```
Slides are created in numerical order.

### Building / GA Images
Must follow the pattern `ANYTHING - Building X - TITLE`:
```
EE_TEMPLE - Building 1 - Midway.jpeg
EE_TEMPLE - Building 1 - Platform 1&2.jpeg
EE_TEMPLE - Building 1 - Ticket Hall.jpeg
```
The **last part after ` - `** becomes the slide title (e.g. `Midway`, `Platform 1&2`).

### 3D Image
Must start with `3D_`:
```
3D_DO93_TEMPLE.png
```
The station name is extracted automatically from this filename.

---

## Output Slide Order

The generated file always follows this order:

1. **Title slide** — station name + 3D image
2. **Design Change Summary** — static content with metadata filled in
3. **2_GA slides** — one per GA drawing found in the P01 input
4. **RF Schematics slides** — one per Design plan image
5. **GA slides** — one per Building image

---

## Stopping the App

Go back to your terminal / command prompt and press:

```
Ctrl + C
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `command not found: python3` | Install Python from [python.org](https://www.python.org/downloads/) and restart your terminal |
| `ModuleNotFoundError: flask` | Run the `pip install` command from Step 2 again |
| `Address already in use` | Another program is using port 5050 — restart your computer and try again |
| Template not found error | Make sure `template_A.pptx` is in the **same folder** as `app.py` |
| Wrong number of GA slides | Use the **GA Count Override** field on the form to manually set the count |
| Building title is wrong | Check the filename follows the `ANYTHING - Building X - TITLE` pattern exactly |

---

## Folder Structure (Reference)

```
APPENDIX_A_AUTOMATION/
├── app.py              ← main application
├── generator.py        ← PPTX generation logic
├── template_A.pptx     ← slide template (do not rename)
├── templates/
│   └── index.html      ← web form
├── uploads/            ← temporary upload storage (auto-created)
└── outputs/            ← generated files (auto-created)
```

---

*For any issues, contact the project maintainer.*
