# Django + React Monorepo

Full-stack monorepo: React (Vite) frontend → **Vercel** · Django API backend → **Railway**.

## Structure

```
.
├── frontend/          # React + Vite app (Vercel)
├── backend/           # Django API (Railway)
│   ├── config/        # Bootstrap & settings only
│   ├── core/          # Shared runtime code
│   └── apps/          # Domain apps (users, …)
├── scripts/           # PowerShell dev helpers
├── docker-compose.yml # Local infrastructure (db)
└── railway.toml       # Railway deployment config
```

---

## Local development

### 1. Start infrastructure

```powershell
docker compose up -d db
```

### 2. Load dev helpers

```powershell
. ./scripts/dev.ps1
```

Available commands:

| Command | Description |
|---|---|
| `Install-BackendDependencies` | Create `.venv` and install `requirements.txt` |
| `Invoke-BackendMigrations` | Run Django migrations |
| `New-BackendSuperUser` | Create a Django superuser |
| `Start-BackendServer` | Start Django dev server on `:8000` |
| `Start-FrontendServer` | Start Vite dev server on `:5173` |
| `Invoke-BackendTests` | Run pytest |

### 3. Environment variables

Copy the example and fill in values:

```powershell
Copy-Item backend/.env.example backend/.env
```

Key variables for local dev:

```env
DJANGO_SETTINGS_MODULE=config.settings.development
SECRET_KEY=change-me
DATABASE_URL=postgres://postgres:postgres@localhost:5432/backend_template
CORS_ALLOWED_ORIGINS=http://localhost:5173
```

---
## Deploy to Railway (backend)

### First deploy

1. **Create a new project** at [railway.app](https://railway.app)
2. **Add a service** → Deploy from GitHub repo → select this repository
3. Railway reads `railway.toml` at the repo root automatically:
   - Builder: Dockerfile (`backend/Dockerfile`)
   - Health check: `GET /health/`
4. **Add a PostgreSQL plugin** inside Railway (generates `DATABASE_URL` automatically)
5. Set the following environment variables in the Railway service:

   | Key | Value |
   |---|---|
   | `DJANGO_SETTINGS_MODULE` | `config.settings.production` |
   | `SECRET_KEY` | A long random string |
   | `ALLOWED_HOSTS` | `.up.railway.app` |
   | `CORS_ALLOWED_ORIGINS` | `https://your-frontend.vercel.app` |
   | `CSRF_TRUSTED_ORIGINS` | `https://your-frontend.vercel.app` |
   | `DJANGO_SECURE_SSL_REDIRECT` | `True` |
   | `DATABASE_URL` | *(injected by PostgreSQL plugin, optional in demo SQLite mode)* |
   | `USE_SQLITE_FOR_DEMO` | `False` (set `True` for temporary demo mode) |

7. Click **Deploy**


## Deploy to Vercel (frontend)

1. **Import** the repository in [vercel.com/new](https://vercel.com/new)
2. Set **Root Directory** to `frontend`
3. Framework preset: **Vite** (auto-detected)
4. Add environment variable:

   | Key | Value |
   |---|---|
   | `VITE_API_URL` | Your Railway backend URL, e.g. `https://your-app.up.railway.app` |

5. Click **Deploy** — Vercel handles build (`npm run build`) and serves `dist/`

> Redeploy automatically on every push to `main` (or configure branch in Vercel settings).

---



### Demo mode: use SQLite on Railway

If you are in demo stage and do not want to provision PostgreSQL yet:

1. In Railway service variables, set `USE_SQLITE_FOR_DEMO=True`
2. Keep `DJANGO_SETTINGS_MODULE=config.settings.production`
3. Redeploy the service

Important: Railway filesystem is ephemeral, so SQLite data can be lost on redeploy/restart. Use PostgreSQL for anything persistent.

### Run migrations on Railway

After the first deploy, open a Railway shell or run via the Railway CLI:

```bash
railway run python backend/manage.py migrate
railway run python backend/manage.py createsuperuser
```

### Railway CLI (optional)

```bash
npm i -g @railway/cli
railway login
railway link          # link local repo to Railway project
railway up            # manual deploy
railway logs          # tail live logs
railway run <cmd>     # run a command inside the Railway environment
```

---

## URLs

| Service | Local | Production |
|---|---|---|
| Frontend | http://localhost:5173 | `https://your-app.vercel.app` |
| Backend API | http://127.0.0.1:8000 | `https://your-app.up.railway.app` |
| Django Admin | http://127.0.0.1:8000/admin/ | `https://your-app.up.railway.app/admin/` |
| API Docs | http://127.0.0.1:8000/api/schema/swagger-ui/ | `https://your-app.up.railway.app/api/schema/swagger-ui/` |
| Health check | http://127.0.0.1:8000/health/ | `https://your-app.up.railway.app/health/` |
