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
- `POST /targets/search` — body `{"type":"contact|project|task|todo|sale_order|opportunity","query":"...","project_id":<id>}`. Returns `{"results":[{id,name,meta}]}`.
- `POST /users/list` — no body needed. Returns `{"users":[{id,partner_id,name,login,email}]}` — the active **internal** Odoo users, for the "Gesendet von" dropdown in the add-in. Cached for 5 minutes.
- `POST /chatter/eml` — attaches the original mail as `.eml` and posts a log note.
- `POST /chatter/note` — posts the mail text as a log note.

## Who appears as the author of the note

Both chatter endpoints accept an optional **`author_id`** (the `partner_id` of the
selected Odoo user, taken from `/users/list`). It is passed straight to Odoo's
`message_post`, so the note shows that person's name and avatar instead of the
technical API user's.

Two things to be aware of:

- The relay **only accepts partner IDs of active internal users**. Without that check
  anyone holding the client token could make a note appear in the name of a customer.
- This is *display* attribution. The technical API user still performs the write, so
  the real writer stays visible in the message's `create_uid` and in the Odoo log.
  Anyone with the client token can pick any name from the list.

If `author_id` is missing, everything behaves exactly as before: the technical API
user is the author.
