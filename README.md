# xlsx-transform

Reads CSV files from `input/` (columns **Page**, **Text**) and creates one Excel file per CSV in `output/`, using the template structure. Input **Page** is written to **Page.** and **Text** to **OE/Owner Comment**.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

### Web UI (upload → download)

1. Ensure the template `output/comment_response_template.xlsx` exists.
2. Start the app:

```bash
python app.py
```

3. Open http://127.0.0.1:5000 in your browser, choose a CSV file, and click **Download Excel** to get the transformed file.

### Command line (batch)

1. Put your CSV file(s) in the `input/` folder (with headers `Page` and `Text`).
2. Ensure the template `output/comment_response_template.xlsx` exists.
3. Run:

```bash
python transform.py
```

Output files are written to `output/` with the same base name as the input (e.g. `input/foo.csv` → `output/foo.xlsx`).

## Input format

- CSV with columns: `Page`, `Text`
- Encoding: UTF-8

## Requirements

- Python 3.10+
- openpyxl, Flask, gunicorn (see `requirements.txt`)

## Deploy on Render (free tier)

Host the web app on [Render](https://render.com) so anyone can use it in the browser.

### Before you deploy

- Push this repo to **GitHub** (or GitLab).
- Keep **`output/comment_response_template.xlsx`** in the repo (do not add it to `.gitignore`), so the app can use it on the server.

### Deploy steps

1. Go to **[render.com](https://render.com)** and sign up or log in (GitHub login is easiest).
2. In the dashboard, click **New +** → **Web Service**.
3. Connect your Git provider and select the **xlsx-transform** repository (or the repo where you pushed this project).
4. Configure the service:
   - **Name:** `xlsx-transform` (or any name you like)
   - **Region:** choose the closest to you
   - **Runtime:** **Python 3**
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT app:app`
   - **Instance type:** **Free** (if available)
5. Click **Create Web Service**. Render will clone the repo, run the build, and start the app.
6. When the build finishes, your app will be live at a URL like **`https://xlsx-transform-xxxx.onrender.com`**. Open it to upload CSV and download Excel.

### Using the Blueprint (optional)

This repo includes **`render.yaml`** (Render Blueprint). After connecting the repo once, you can use **New + → Blueprint** and point it at this repo; Render will create the Web Service from the YAML so you don’t have to fill in build/start commands manually.

---

### Other platforms

- **Railway:** New Project → Deploy from GitHub → select this repo. It will use the `Procfile`. Add a public domain in the service settings.
- **Heroku:** Create an app, connect the repo. The `Procfile` and `runtime.txt` are already set up.
