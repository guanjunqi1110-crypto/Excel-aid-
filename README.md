# Excel Readiness AI Coach

Streamlit prototype that diagnoses Excel and business workflow readiness for new accounting and finance students, using the bundled **Excel_Readiness_AI_Coach_Content_Pack.xlsx**.

**Project folder on this PC:** `C:\Users\guanj\excel-readiness-ai-coach\`

## Run locally

```powershell
cd C:\Users\guanj\excel-readiness-ai-coach
pip install -r requirements.txt
streamlit run app.py
```

## Put it on the internet (Streamlit Community Cloud)

This machine may not have **Git** in the path; install from [git-scm.com](https://git-scm.com/download/win) if `git` is not recognized.

### Step 1 — Create an empty GitHub repository

1. Log in to [github.com](https://github.com) and click **New repository**.
2. Name it (e.g. `excel-readiness-ai-coach`).  
3. Choose **Public**.  
4. **Do not** add a README, `.gitignore`, or license (keeps the first push simple).  
5. Create the repo and copy the HTTPS URL, e.g. `https://github.com/YourUsername/excel-readiness-ai-coach.git`.

### Step 2 — Push this folder to GitHub

In **PowerShell** (Run as you normally do):

```powershell
cd C:\Users\guanj\excel-readiness-ai-coach
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\push-to-github.ps1 -GitHubUrl "https://github.com/YourUsername/excel-readiness-ai-coach.git"
```

(Replace with your real repo URL.) If Git isn’t installed, install it and run the script again.

**Manual alternative:** `git init`, `git add`, `git commit`, `git remote add origin <url>`, `git push -u origin main`.

### Step 3 — Deploy on Streamlit Cloud

1. Open **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub.  
2. **New app** → select your new repository, branch `main`, main file **`app.py`**.  
3. **Deploy.** Your app will get a public URL like `https://your-app-name.streamlit.app` — share that with your professor.

### Step 4 (optional) — OpenAI on the server

In the Streamlit Cloud app: **⋮ (Manage app) → Settings → Secrets** and add:

```toml
OPENAI_API_KEY = "sk-...your-key..."
```

Redeploy or restart the app. **Do not** commit API keys to GitHub.

## Local OpenAI (optional, for `streamlit run` on your PC)

Create `.streamlit/secrets.toml` (this file is gitignored):

```toml
OPENAI_API_KEY = "sk-..."
```

This project is for classroom use (demo / prototype).
