"""
Outlook -> Odoo Relay (Railway / FastAPI)

Purpose-built gateway between the Outlook add-in and Odoo. It is NOT a generic
proxy: it exposes only the specific operations the add-in needs. The Odoo API
key lives here (as a Railway environment variable) and never reaches the browser.

Configuration comes entirely from environment variables (set in Railway):
  ODOO_BASE_URL   e.g. https://dynaplo.odoo.com   (no trailing slash)
  ODOO_API_KEY    the Odoo API key (secret)
  ODOO_DB         database name; for Odoo Online usually the subdomain (optional)
  CLIENT_TOKEN    shared secret the add-in must send in the X-Client-Token header
  ALLOWED_ORIGIN  the GitHub Pages origin allowed to call this relay
"""

import os
import re
from typing import Any, Optional

import httpx
from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# --- Konfiguration (in Railway setzen) ---
ODOO_BASE_URL = os.environ.get("ODOO_BASE_URL", "").rstrip("/")
ODOO_API_KEY = os.environ.get("ODOO_API_KEY", "")
ODOO_DB = os.environ.get("ODOO_DB", "")
CLIENT_TOKEN = os.environ.get("CLIENT_TOKEN", "")
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://nicdyna.github.io")

app = FastAPI(title="Outlook -> Odoo Relay", version="0.1.0")

# Nur die GitHub-Pages-Herkunft darf das Relay aus dem Browser aufrufen.
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "X-Client-Token"],
)


def _check_token(token: Optional[str]) -> None:
    """Nur Aufrufe mit dem korrekten Client-Token zulassen."""
    if not CLIENT_TOKEN or token != CLIENT_TOKEN:
        raise HTTPException(status_code=401, detail="Ungültiger oder fehlender Client-Token.")


def _odoo_headers() -> dict:
    headers = {
        "Authorization": f"bearer {ODOO_API_KEY}",
        "Content-Type": "application/json; charset=utf-8",
    }
    if ODOO_DB:
        headers["X-Odoo-Database"] = ODOO_DB
    return headers


async def _odoo_call(model: str, method: str, payload: dict) -> Any:
    """Ruft eine Odoo-Methode über die JSON-2-API auf und gibt das Roh-Ergebnis zurück."""
    if not ODOO_BASE_URL or not ODOO_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="Relay nicht konfiguriert (ODOO_BASE_URL / ODOO_API_KEY fehlen).",
        )
    url = f"{ODOO_BASE_URL}/json/2/{model}/{method}"
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(url, headers=_odoo_headers(), json=payload)
    except httpx.RequestError as exc:
        raise HTTPException(status_code=504, detail=f"Odoo nicht erreichbar: {exc}")

    if resp.status_code != 200:
        # Odoo-Fehler weiterreichen (gekürzt; enthält keine Secrets)
        raise HTTPException(
            status_code=502,
            detail=f"Odoo-Fehler ({resp.status_code}): {resp.text[:500]}",
        )
    return resp.json()


def _company_of(record: dict) -> str:
    """Firmenname robust ermitteln, unabhängig davon, wie Odoo many2one serialisiert."""
    parent = record.get("parent_id")
    if isinstance(parent, list) and len(parent) >= 2:
        return parent[1] or ""
    if isinstance(parent, dict):
        return parent.get("display_name") or parent.get("name") or ""
    ccn = record.get("commercial_company_name") or ""
    name = record.get("name") or ""
    return ccn if ccn and ccn != name else ""


def _extract_id(result: Any) -> Optional[int]:
    """Neue Datensatz-ID robust aus der Odoo-Antwort ziehen (int / [int] / [{id}] / {id})."""
    if isinstance(result, int):
        return result
    if isinstance(result, dict):
        return result.get("id")
    if isinstance(result, list) and result:
        first = result[0]
        if isinstance(first, int):
            return first
        if isinstance(first, dict):
            return first.get("id")
    return None


def _safe_filename(name: str) -> str:
    name = (name or "E-Mail.eml").strip()
    name = re.sub(r'[\/\\:*?"<>|]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    if not name.lower().endswith(".eml"):
        name += ".eml"
    return name[:120] or "E-Mail.eml"


def _html_escape(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class PartnerSearch(BaseModel):
    query: str


class EmlAttach(BaseModel):
    partner_id: int
    filename: str
    eml_base64: str
    subject: str = ""


@app.get("/health")
async def health():
    """Schneller Funktionstest (kein Token nötig)."""
    return {
        "status": "ok",
        "configured": bool(ODOO_BASE_URL and ODOO_API_KEY and CLIENT_TOKEN),
    }


@app.post("/partners/search")
async def partners_search(
    body: PartnerSearch,
    x_client_token: Optional[str] = Header(default=None),
):
    """Sucht Kontakte (res.partner) nach Name oder E-Mail (Teiltext)."""
    _check_token(x_client_token)

    query = (body.query or "").strip()
    if len(query) < 2:
        return {"partners": []}

    payload = {
        "domain": ["|", ["name", "ilike", query], ["email", "ilike", query]],
        "fields": ["name", "email", "parent_id", "commercial_company_name", "is_company"],
        "limit": 20,
        "order": "name asc",
    }
    result = await _odoo_call("res.partner", "search_read", payload)

    partners = []
    for r in (result or []):
        partners.append({
            "id": r.get("id"),
            "name": r.get("name") or "",
            "email": r.get("email") or "",
            "company": _company_of(r),
            "is_company": bool(r.get("is_company")),
        })
    return {"partners": partners}


@app.post("/chatter/eml")
async def chatter_eml(
    body: EmlAttach,
    x_client_token: Optional[str] = Header(default=None),
):
    """Hängt die Original-E-Mail als .eml-Datei an die Chatter eines Kontakts an."""
    _check_token(x_client_token)

    filename = _safe_filename(body.filename)

    # 1) Anhang direkt am Kontakt anlegen
    attachment_vals = {
        "name": filename,
        "datas": body.eml_base64,
        "mimetype": "message/rfc822",
        "res_model": "res.partner",
        "res_id": body.partner_id,
    }
    create_result = await _odoo_call("ir.attachment", "create", {"vals_list": [attachment_vals]})
    attachment_id = _extract_id(create_result)
    if not attachment_id:
        raise HTTPException(
            status_code=502,
            detail=f"Anhang-ID nicht erkannt. Odoo-Antwort auf create: {create_result}",
        )

    # 2) Interne Chatter-Notiz mit verknüpftem Anhang
    subject = body.subject or filename
    note_body = f"<p>E-Mail archiviert: {_html_escape(subject)}</p>"
    post_result = await _odoo_call("res.partner", "message_post", {
        "ids": [body.partner_id],
        "body": note_body,
        "message_type": "comment",
        "subtype_xmlid": "mail.mt_note",
        "attachment_ids": [attachment_id],
    })

    return {
        "ok": True,
        "attachment_id": attachment_id,
        "message_id": _extract_id(post_result),
        "partner_url": f"{ODOO_BASE_URL}/web#id={body.partner_id}&model=res.partner&view_type=form",
    }
