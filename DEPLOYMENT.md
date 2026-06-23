# Deployment Guide — Growth Stack Talent Candidate Portal

## Running locally

```bash
cd candidate-form
python3 -m venv venv
venv/bin/pip install -r requirements.txt
cp .env.example .env        # then edit .env with your values
venv/bin/python app.py
```

Open http://localhost:5050 — the candidate form.  
Open http://localhost:5050/admin — the admin portal.

---

## Deploying to Render (recommended — free tier available)

Render is the simplest way to put the form online with a public URL.

### 1. Push to GitHub

```bash
cd candidate-form
git init
echo "venv/" >> .gitignore
echo ".env" >> .gitignore
echo "candidates.csv" >> .gitignore
echo "output/" >> .gitignore
echo "uploads/" >> .gitignore
git add .
git commit -m "Initial commit"
```

Create a new GitHub repo and push to it.

### 2. Create a Render Web Service

1. Go to https://render.com and sign in.
2. Click **New → Web Service**.
3. Connect your GitHub repo.
4. Set the following:

| Setting | Value |
|---|---|
| **Environment** | Python 3 |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn app:app` |

### 3. Set environment variables in Render

In your service's **Environment** tab, add:

| Key | Value |
|---|---|
| `FLASK_SECRET_KEY` | A long random string |
| `ADMIN_PASSWORD` | Your chosen admin password |
| `SMTP_HOST` | e.g. `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | Your email address |
| `SMTP_PASSWORD` | Your email app password |
| `RECRUITER_EMAIL` | Where notification emails go |

> **Gmail users:** You must use an [App Password](https://myaccount.google.com/apppasswords), not your regular Gmail password. Enable 2FA on your Google account first, then generate an App Password under Security → App Passwords.

### 4. Persistent storage (important)

Render's free tier has an **ephemeral filesystem** — files written to disk (candidates.csv, uploads/, output/) are lost on each deploy or restart.

To keep data permanently, add a **Render Disk**:
1. In your service settings, go to **Disks → Add Disk**.
2. Mount path: `/data`
3. Update `app.py` to point storage at `/data`:

```python
BASE_DIR = Path(os.environ.get("DATA_DIR", Path(__file__).parent))
```

Add `DATA_DIR=/data` to your Render environment variables.

---

## Deploying to Railway (alternative)

1. Go to https://railway.app and sign in.
2. Click **New Project → Deploy from GitHub repo**.
3. Select your repo — Railway auto-detects the Procfile.
4. Add the same environment variables as above under **Variables**.
5. Add a **Volume** (under your service settings) mounted at `/data` for persistent storage, then set `DATA_DIR=/data`.

---

## Email setup (Gmail step-by-step)

1. Enable 2-Step Verification on your Google account.
2. Go to https://myaccount.google.com/apppasswords.
3. Create an app password named "Growth Stack Talent Portal".
4. Use the generated 16-character password as `SMTP_PASSWORD`.
5. Set `SMTP_HOST=smtp.gmail.com`, `SMTP_PORT=587`, `SMTP_USER=your@gmail.com`.

---

## Admin portal

- URL: `https://your-app-url/admin`
- Password: whatever you set as `ADMIN_PASSWORD`
- Shows all submissions in a table (most recent first)
- Download buttons for each candidate's Word front sheet and uploaded CV

---

## File structure

```
candidate-form/
├── app.py                  # Flask application
├── requirements.txt        # Python dependencies
├── Procfile                # Gunicorn start command
├── .env                    # Your local environment variables (never commit)
├── .env.example            # Template for environment variables
├── static/
│   ├── logo.png            # Company logo
│   └── style.css           # Styles
├── templates/
│   ├── form.html           # Candidate application form
│   ├── admin_login.html    # Admin login page
│   └── admin_dashboard.html # Submissions viewer
├── candidates.csv          # CRM/ATS data store
├── output/                 # Generated Word front sheets
└── uploads/                # Uploaded CVs
```
