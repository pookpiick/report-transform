# report-transform

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

## Deploy on Render

Use this section when you want to (re)deploy the web app on [Render](https://render.com).

### Before you deploy

- Push this repo to **GitHub** (or GitLab).
- Keep **`output/comment_response_template.xlsx`** in the repo (do not add it to `.gitignore`).
- **If `git push` fails with "Permission denied (publickey)":** use HTTPS and a token:
  1. GitHub → **Settings → Developer settings → Personal access tokens** → create a token with `repo` scope.
  2. `git remote set-url origin https://github.com/YOUR_USERNAME/report-transform.git`
  3. Push again; when prompted for password, paste the **token**.

### Deploy steps

1. Go to **[render.com](https://render.com)** and sign in (e.g. with GitHub).
2. Click **New +** → **Web Service** (the “Dynamic web app…” option).
3. Connect your Git provider and select the **report-transform** repo → **Connect**.
4. Configure the service:

   | Field | Value |
   |-------|--------|
   | Name | `report-transform` (or any name) |
   | Region | Closest to you |
   | Runtime | **Python 3** |
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn --bind 0.0.0.0:$PORT app:app` |
   | Instance type | **Free** (if available) |

5. Click **Create Web Service**. Wait until the service is **Live**.
6. Open the URL (e.g. `https://report-transform-xxxx.onrender.com`) to use the app.

### Blueprint (optional)

This repo has **`render.yaml`**. After the repo is connected, you can use **New + → Blueprint** and select this repo; Render will create the Web Service from the YAML so you don’t have to enter build/start commands.

---

### Other platforms

- **Railway:** New Project → Deploy from GitHub → select this repo. It will use the `Procfile`. Add a public domain in the service settings.
- **Heroku:** Create an app, connect the repo. The `Procfile` and `runtime.txt` are already set up.
