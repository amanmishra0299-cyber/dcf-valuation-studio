# Deploying from a phone

You don't need a computer for any of this. Everything below is a browser on your phone.

Target: a permanent `https://` link that opens in Chrome for you and anyone you send it to.

## Path A — Render free tier (recommended)

Render auto-detects `render.yaml`, so there is almost nothing to configure.

**1. Put the code on GitHub** (~5 min)
- Go to https://github.com → **Sign up** (email + password).
- Tap **+** (top right) → **New repository** → name it `dcf-valuation-studio` → **Create repository**.
- On the empty repo page tap **uploading an existing file**.
- Upload every file from the zip **keeping the `engine/` folder**:
  `app.py`, `requirements.txt`, `Procfile`, `runtime.txt`, `render.yaml`, `fly.toml`,
  `.gitignore`, `README.md`, `tests.py`, `ui_test.js`,
  `templates/index.html`, and the five files inside `engine/`.
  GitHub's web uploader lets you create the `engine/` and `templates/` folders by typing
  the name with a slash, e.g. `engine/valuation.py`.
- Tap **Commit changes**.

**2. Create the service on Render** (~3 min)
- Go to https://render.com → **Get Started** → sign up **with GitHub** (saves re-entering details).
- Dashboard → **New +** → **Web Service** → **Connect a GitHub account** → allow Render →
  pick `dcf-valuation-studio`.
- Instance type: **Free**. Everything else is already filled in from `render.yaml`.
- Tap **Create Web Service**.

**3. Wait ~3–4 minutes** for the build. When the log says `Listening at`, the URL at the
top of the page (something like `https://dcf-valuation-studio.onrender.com`) is your link.
Open it in Chrome. Done.

### Two things to expect on the free tier
- **Cold start.** After ~15 minutes unused the container sleeps and the first tap takes
  ~50 seconds to wake it. Tapping again is instant.
- **Screener may block cloud IPs.** screener.in throttles datacentre traffic harder than
  home traffic. If Analyse returns "Could not find 'X'" for a symbol you know is valid,
  that's a block, not a bug — the **Import Excel** and **manual entry** tabs work
  regardless, because they don't call screener at all.

## Path B — Railway

https://railway.app → sign in with GitHub → **New Project** → **Deploy from GitHub repo**.
It reads `Procfile` automatically. No free tier any more (roughly $5 of credit to start),
but no cold starts and fewer IP blocks than Render free.

## Path C — Fly.io

`fly.toml` is included. This one needs the `flyctl` command-line tool, so it isn't
phone-friendly. Only use it if you have a computer later.

## What each file is for

| File | Purpose |
|---|---|
| `render.yaml` | Render's build + start command + health check path |
| `Procfile` | The same start command for Railway/Heroku-style platforms |
| `runtime.txt` | Pins Python 3.12 |
| `requirements.txt` | Dependencies, including `gunicorn` for production serving |
| `fly.toml` | Fly.io config |

The app already reads `PORT` from the environment and binds `0.0.0.0`, which is what all
three platforms need — no code change required.

## Verifying a deployment

Open these two URLs on the deployed app:

- `/api/health` → should return `{"cache":0,"ok":true,...}`
- `/` → the app itself

If `/api/health` works but Analyse fails, the deploy is fine and screener.in is blocking
the host — use manual entry, or move to a paid/residential-egress host.
