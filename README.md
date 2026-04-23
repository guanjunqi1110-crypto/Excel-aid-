# Excel Readiness AI Coach

Streamlit prototype that diagnoses Excel and business workflow readiness for new accounting and finance students, using the bundled **`Excel_Readiness_AI_Coach_Content_Pack_with_books.xlsx`** (older content-pack workbooks in the repo are optional/legacy).

**Project folder on this PC (may vary):** e.g. `C:\Users\guanj\Downloads\` (Streamlit `app.py` and content pack in the same folder, or a dedicated project folder)

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

### Step 4 (optional) — OpenAI so visitors never type a key

**Never** put a real `sk-` key in `app.py` or any file committed to a **public** repo.

- **Streamlit Cloud (the public site):** set the key **once** in **Manage app → Settings → Secrets**:

```toml
OPENAI_API_KEY = "sk-...your-key..."
```

Save → **Reboot** the app. Then everyone who opens your `.streamlit.app` link gets live AI without pasting (API usage is billed to your OpenAI account—set limits at platform.openai.com).

- **Local `streamlit run` only:** copy `openai_key_local.txt.example` to `openai_key_local.txt`, one line with your `sk-` key. That file is **gitignored** and never goes to GitHub.

## Local OpenAI (optional, for `streamlit run` on your PC)

Create `.streamlit/secrets.toml` (this file is gitignored):

```toml
OPENAI_API_KEY = "sk-..."
```

This project is for classroom use (demo / prototype).
