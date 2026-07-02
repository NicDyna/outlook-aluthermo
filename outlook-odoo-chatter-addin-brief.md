# Outlook → Odoo Chatter Add-in — Project Brief

**Date:** 2026-07-02 · **Revised:** 2026-07-02 (all major decisions locked)
**Owner:** Nicola (Dynaplo) — not a developer, works step-by-step, commits via GitHub Desktop only
**Target:** Odoo 19 Enterprise, Odoo Online (SaaS)

**Goal:** An Outlook add-in that, for the currently open email, either
(A) posts a **clean text version** of the mail to the chatter of a contact (`res.partner`) as a private log note, or
(B) attaches the full original **`.eml` file** to that contact's chatter.

---

## 1. What we are building

A single **Office Web Add-in** (JavaScript + XML manifest, static files only) that runs in:

- **Classic Outlook on Windows** (primary target — renders add-ins in an embedded Edge WebView2)
- Outlook on the web (outlook.office.com)
- New Outlook on Windows

One codebase, one manifest, all three clients. Add-ins are installed per **mailbox**, not per machine: sideloading once via Outlook on the web makes the add-in appear in Classic Outlook automatically (same M365 account).

**Two supported systems behind the add-in:** a **Railway** back-end (the Odoo gateway) and **n8n** (the email-cleaning brain, used only on the text path). See §3.

**Explicitly NOT building:** a COM/VSTO add-in (compiled .NET, per-machine installers, dead end — New Outlook doesn't support COM).

---

## 2. The user flow

1. Customer opens a mail in Outlook and clicks the ribbon button **"An Odoo senden"** → a task pane (docked side panel) opens.
2. The task pane asks:
   - **How to send to Odoo:** `.eml` file **or** text in the chatter log note.
   - **If text:** paste **everything** (full thread) **or** **only the last message**.
   - **Which contact** to attach it to (search all `res.partner`, company shown per row; see §6.1).
   - **Done** / **Cancel** (review step before anything is written).
3. `.eml` path → attach the raw file to the contact's chatter (no processing).
4. Text path → body goes to n8n, is cleaned, comes back tidy, and lands in the contact's chatter as a **private internal note**.

Both actions post as a **private internal log note** (`mail.mt_note`): visible to the team inside Odoo, **never** emails the contact or notifies followers.

Works on **received and sent** mail. Received mail matches on the sender; Sent Items also offers the To/Cc recipients as the contact to match.

---

## 3. Architecture

```
Add-in (static JS on GitHub Pages, German UI, no secrets)
   │  fetch()  — CORS locked to the GitHub Pages origin + a client token
   ▼
Railway FastAPI relay  ── holds the Odoo API key + instance URL (env vars)
   │                    ── assembles the final note + runs the text-integrity check
   │  ├── POST /partners/search ─┐
   │  ├── POST /chatter/eml ──────┼──►  https://<instance>.odoo.com/json/2/...  (Bearer key)
   │  └── POST /chatter/note ─────┘        server-to-server, no CORS problem
   │            │  (text path only)
   │            ▼
   │         n8n  ──►  Claude   (returns WHERE to cut, never the text itself)
```

- **Frontend:** pure static HTML/CSS/JS on **GitHub Pages** (HTTPS, which Office add-ins require). No build step, no framework — plain JS keeps the GitHub-Desktop workflow simple.
- **Railway relay (single Odoo front door):** a small FastAPI app. It is **purpose-built, not a generic proxy** — it exposes only the three operations we need, so even if the URL leaked, it could not run arbitrary Odoo commands. It holds the Odoo API key server-side (env var), so the key never touches the browser. It also assembles the note and performs the verbatim text check (§5).
- **n8n (email-analysis brain):** called by Railway only on the text path. Holds the Claude prompt + Claude API key. Returns cut points for the body. Being visual, the cleaning prompt can be tuned by Nicola without editing code.
- **Why both:** Railway = the one trusted place for Odoo credentials + the integrity guarantee; n8n = the editable AI pipeline. One place owns the Odoo key; one place owns the AI prompt.

### 3.1 Secrets & relay protection

- **Odoo API key:** stored **only** as a Railway environment variable. Never committed, never in the add-in, never in the browser.
- **Claude API key:** stored **only** in n8n. Never in the add-in.
- **No settings screen for credentials** in the add-in (the roamingSettings key screen from the original plan is dropped — the key lives on Railway instead). The instance URL used for deep-links (§6.4) is public and may be a plain config constant in the add-in JS.
- **Protecting the relay:** restrict `Access-Control-Allow-Origin` to the GitHub Pages origin, and require a **client token** the add-in sends on every call.
- **Honest caveat:** the add-in's JS is public (GitHub Pages), so a client token baked into it is readable by anyone who views the source. Combined with the CORS lock and the relay's three narrow actions on one Odoo instance, real-world risk is low. **Future hardening (not v1):** validate the Outlook user's Microsoft identity token in the relay so only your tenant's authenticated users can call it.

---

## 4. Verified platform facts (checked 2026-07-02)

1. **Full .eml is available in JavaScript:** `Office.context.mailbox.item.getAsFileAsync(callback)` returns the current message as a **Base64-encoded EML string**. API set **Mailbox 1.14**, minimum permission `read item`, **Message Read mode only** (does not work on drafts/compose).
   Docs: https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread
2. Microsoft's own docs direct developers to `getAsFileAsync` specifically **in Classic Outlook on Windows** when Base64 .eml is needed — so Classic desktop (current M365 builds) supports it.
3. **Clean text body:** `item.body.getAsync(Office.CoercionType.Text, cb)` — available since the earliest API versions, works everywhere.
4. **Header metadata is structured and directly available** (no parsing needed): `item.from`, `item.to`, `item.cc`, `item.subject`, `item.dateTimeCreated`. These build the note header and are **never touched by AI** (§5).
5. Known minor quirk: the .eml generated by Classic desktop can have slightly different headers (e.g. Message-ID) than the server copy. Irrelevant for archiving; do not build automation on those headers.
6. The taskpane is a webpage inside WebView2 → **normal browser CORS rules apply.** Direct `fetch()` to `*.odoo.com` is blocked; all Odoo calls go through the Railway relay.
7. Gotcha: the Read item context requires the **Reading Pane to be on** (or the message opened in its own window) in Classic Outlook.

**Runtime feature detection (required):**

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.14')) {
  // enable the ".eml attach" button
} else {
  // disable it with a German tooltip explaining the client is too old
}
```

Keep the manifest's minimum requirement set LOW (e.g. Mailbox 1.5) so the add-in installs broadly; gate the .eml feature at runtime.

---

## 5. Text-note cleaning + the text-integrity guarantee

The note has two parts, handled very differently:

**Part 1 — the header (deterministic, never AI).** Built by us from Outlook's structured fields: sender name + email, CC, date, subject. No chance of a wrong name or address.

```
Von: Max Mustermann <max@example.com>
CC: ...
Datum: 2026-07-02 14:31
Betreff: ...
```

**Part 2 — the body (the only hard part).** Strip signatures; separate the latest message from earlier ones. AI-assisted (decision locked), used **only to locate the cut points** — never to produce or rewrite text.

### The guarantee (Nicola's non-negotiable: email text must never be altered)

1. n8n sends the plain-text body to Claude.
2. Claude returns **where to cut** (boundaries: signature start, "older message begins here"), never rewritten text.
3. The Railway relay does the actual slicing on the **original** body string.
4. Before posting, the relay **verifies the cleaned text appears verbatim** (exact substring) in the original. If it does not match character-for-character, the AI result is rejected and we fall back to the raw body (or a rules-based cut).

⇒ The words posted to Odoo are always sliced straight from the real email and provably unchanged. Claude only decides *where the scissors go*.

### Mode behaviour

- **"Only the last message":** header + latest message, signature stripped.
- **"Everything":** header + latest message (signature stripped) + separator `─── Vorherige Nachricht ───` + the earlier thread history **passed through as-is** (not re-cut, to keep the integrity check simple and preserve context).

German angle: quoting headers are German ("Von: / Gesendet: / Am … schrieb …"); AI handles this variety better than English-centric rule parsers.

---

## 6. Odoo integration (Odoo 19 SaaS, JSON-2 API)

All Odoo calls run **server-side from the Railway relay** using the **JSON-2 external API** with **Bearer token auth** — the same pattern proven in Nicola's Firefox timesheet extension and n8n geocoding workflow (`POST https://<instance>.odoo.com/json/2/<model>/<method>`, `Content-Type: application/json`, `Authorization: Bearer <key>`). **Reuse that exact request shape;** verify the current body format (ids/kwargs) against the Odoo 19 docs and the open-source repo (github.com/odoo/odoo, branch 19.0) before coding — do not guess.

### 6.1 Find the contact — relay `POST /partners/search`

1. Add-in reads sender via `item.from` (received) or offers To/Cc (Sent Items).
2. Relay searches `res.partner` by exact email first (`email =ilike sender`), then falls back to a **searchable combobox** (name/email substring). Reuse the combobox pattern from the Firefox timesheet extension.
3. Search **all** partners (people and companies). Each result row shows **name, email, company**. User confirms one. **Never auto-post without confirmation.**

### 6.2 Action A — clean text to chatter — relay `POST /chatter/note`

1. Add-in sends the structured metadata (from/to/cc/subject/date) + raw text body + mode (`full` | `last`) + confirmed `partner_id`.
2. Relay: (text path) call n8n → get cut points → slice + verify verbatim (§5) → build HTML note.
3. Optionally append an attachment-filename line: `Anhänge: file1.pdf, file2.xlsx` (decision: include it).
4. HTML-escape the text body before embedding (`<`, `>`, `&`).
5. Post as an **internal log note**: `message_post` on `res.partner` with `body=<html>` and `subtype_xmlid='mail.mt_note'`. Verify the exact `message_post` signature (kwarg names, JSON-2 callability) in `addons/mail/models/mail_thread.py` on branch 19.0.

### 6.3 Action B — attach full .eml to chatter — relay `POST /chatter/eml`

1. Add-in `item.getAsFileAsync` → Base64 EML string (passes straight into `ir.attachment.datas`, which expects Base64).
2. Relay creates the attachment:

```json
POST /json/2/ir.attachment/create
{
  "name": "2026-07-02_Betreff-sanitized.eml",
  "datas": "<base64 EML>",
  "mimetype": "message/rfc822",
  "res_model": "res.partner",
  "res_id": <partner_id>
}
```

3. Post a chatter note linking it: `message_post` with `attachment_ids=[<attachment_id>]` and a short body ("E-Mail archiviert: <subject>").
4. **Verify attachment-linking rules** in `mail_thread.py` (`_message_post_process_attachments` / access checks): `message_post` only links attachments the API user owns that are either already on the target record or parked on `res_model='mail.compose.message', res_id=0`. If creating directly on `res.partner` fails the check, switch to the `mail.compose.message` parking pattern.
5. Filename: `YYYY-MM-DD_<subject>.eml`, subject sanitized (strip `/\:*?"<>|`, collapse whitespace, cap ~80 chars).
6. Size: large-attachment emails become large Base64 JSON bodies. Test with a ~15–25 MB mail early; if Odoo SaaS or Railway rejects it, cap with a clear German error message.

### 6.4 Deep link back

Success state shows the created record as a link: `https://<instance>.odoo.com/odoo/contacts/<id>` — verify URL scheme in Odoo 19.

### 6.5 General Odoo rules (standing agreement)

- Verify every model/method/field against the open-source GitHub repo (19.0) before use.
- **No destructive or bulk operations.** Only ever creates messages/attachments on a single, user-confirmed record per click.
- API user = whoever owns the API key; chatter author will be that user. Acceptable for v1.
- **Future targets (out of scope v1):** keep the Odoo model name a config constant so `crm.lead` / `helpdesk.ticket` / `sale.order` can be added later as a one-line change.

---

## 7. Add-in specifics

- **Manifest:** classic **XML manifest** (universally supported, simplest for static hosting; skip the unified JSON manifest for now).
  - `<Permissions>ReadItem</Permissions>`.
  - Requirement set minimum: Mailbox 1.5 (runtime-gate 1.14 for the .eml button, see §4).
  - Activation: Message Read surface; ribbon button opening a **taskpane** (label: "An Odoo senden").
  - Icons at 16/32/80 px (also provide 64/128), PNG, hosted on the same GitHub Pages site, referenced by absolute HTTPS URLs.
  - All URLs in the manifest must be absolute HTTPS URLs to the Pages site.
- **UI language: German.** Suggested labels: "An Odoo senden", "Kontakt suchen…", "Nur Text in Chatter", "Komplette E-Mail (.eml) anhängen", "Alles" / "Nur letzte Nachricht", "Überprüfen & senden", "Erfolgreich gesendet ✓".
- **Cache busting:** WebView2 caches aggressively. Reference `taskpane.js?v=<date>` and bump the version on every deploy; document a hard-reload fallback (close/reopen taskpane, or clear Office web cache).

---

## 8. Repo layout (GitHub Pages from repo root or /docs)

```
outlook-odoo-chatter/
├── manifest.xml
├── index.html            (taskpane)
├── taskpane.js
├── styles.css
├── assets/
│   ├── icon-16.png … icon-128.png
├── relay/                (Railway FastAPI app — deployed separately, holds Odoo key)
│   ├── main.py
│   └── requirements.txt
├── n8n/
│   └── clean-workflow.json   (exported n8n workflow, for reference/versioning)
└── README.md             (setup + sideload instructions in simple steps)
```

---

## 9. Decisions — locked

1. **Relay:** Railway FastAPI (server-to-server to Odoo, no CORS).
2. **API key location:** server-side on Railway (add-in never sees it).
3. **Post visibility:** private internal note (`mail.mt_note`).
4. **Partner scope:** search all contacts; show company per row.
5. **Sent Items:** supported; recipient (To/Cc) matching in addition to sender.
6. **Text sub-option:** full thread vs. only the last message.
7. **Body cleaning:** AI-assisted via n8n + Claude, using the "AI picks cut points, text sliced from original and verified verbatim" guarantee (§5).
8. **Attachment filenames** appended to the text note.
9. **Future targets** (crm.lead, etc.) kept as a config constant, not built now.

---

## 10. Workflow rules for Claude Code (non-negotiable)

1. Claude Code edits files locally; **Nicola commits and pushes via GitHub Desktop** — never run `git` commands.
2. GitHub Pages auto-deploys on push to the configured branch.
3. Explain every step in plain language; Nicola is not a developer.
4. Verify Odoo internals against github.com/odoo/odoo (19.0) instead of assuming.
5. Never place the API key or instance-specific secrets in committed files.

---

## 11. Sideloading & debugging

- **Install once via Outlook on the web:** open https://aka.ms/olksideload → My add-ins → "Add a custom add-in" → "Add from URL" (the raw GitHub Pages URL of `manifest.xml`). Because add-ins are mailbox-scoped, it then appears in Classic Outlook desktop automatically (restart Outlook).
- If the earlier Outlook-Online prototype add-in is still sideloaded, it may already be visible in Classic Outlook — check before assuming anything is broken.
- **Debug in OWA first** (F12 DevTools, same code), only smoke-test in Classic Outlook afterwards.
- Classic Outlook check: reading pane on; File → Office Account to confirm an up-to-date M365 build (needed for Mailbox 1.14).

---

## 12. Milestones

- **M1:** Repo + manifest + empty German taskpane; sideloaded via OWA; visible and opening in Classic Outlook.
- **M2:** Taskpane shows subject, sender, CC, date, plain-text body of the open mail; the choice UI (eml/text, full/last, contact) laid out (no back-end yet).
- **M3:** Railway relay deployed; `/partners/search` working end-to-end from the taskpane.
- **M4:** Action B (.eml) live — attachment + note posted; test with a large email.
- **M5:** n8n cleaning workflow + Railway `/chatter/note`; Action A (text) live with the verbatim integrity check; verify "full" vs "last" modes in Odoo.
- **M6:** Review step, error states, cache-busting, README with setup steps; relay client-token + CORS lock in place.

---

## 13. Verification checklist before writing integration code

- [ ] Exact JSON-2 request/response shape for `create`, `search`-style and generic method calls (Odoo 19 docs + repo) — cross-check with the working Firefox-extension calls.
- [ ] `message_post` kwargs on 19.0 (`body`, `subtype_xmlid`, `attachment_ids`) and its attachment access rules.
- [ ] Is `search_read` exposed via JSON-2, or use `search` + `read`?
- [ ] Odoo 19 web URL scheme for deep-linking a contact.
- [ ] Mailbox 1.14 actually reported by `isSetSupported` in Nicola's Classic Outlook build.
- [ ] n8n webhook contract: request (body + mode) and response (cut points) shape.
- [ ] Claude prompt returns boundaries only; verbatim check rejects any non-matching slice.
