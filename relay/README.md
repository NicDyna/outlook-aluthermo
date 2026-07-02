# Relay (Railway) — Outlook → Odoo

Small FastAPI service. It is the single, **purpose-built** gateway between the
Outlook add-in and Odoo — it exposes only the operations the add-in needs, so it
cannot be used to run arbitrary Odoo commands. The Odoo API key lives **here**
(as a Railway environment variable) and never reaches the browser or the repo.

## Environment variables (Railway → your service → Variables)

| Variable        | Example                    | Purpose |
|-----------------|----------------------------|---------|
| `ODOO_BASE_URL` | `https://dynaplo.odoo.com` | Your Odoo instance (no trailing slash) |
| `ODOO_API_KEY`  | *(secret)*                 | Odoo API key — create in Odoo, paste here only |
| `ODOO_DB`       | `dynaplo`                  | Database name; for Odoo Online usually the subdomain |
| `CLIENT_TOKEN`  | *(random string)*          | Shared secret the add-in must send (`X-Client-Token`) |
| `ALLOWED_ORIGIN`| `https://nicdyna.github.io`| GitHub Pages origin allowed to call the relay |

## Railway settings
- Set the service **Root Directory** to `relay`.
- Build: Railway (Nixpacks) auto-installs `requirements.txt`.
- Start: taken from the `Procfile` → `uvicorn main:app --host 0.0.0.0 --port $PORT`.

## Endpoints
- `GET /health` — quick check, no token needed. Returns `{"status":"ok","configured":true}` once all variables are set.
- `POST /partners/search` — body `{"query":"..."}`, header `X-Client-Token: <CLIENT_TOKEN>`. Returns `{"partners":[{id,name,email,company,is_company}]}`.
