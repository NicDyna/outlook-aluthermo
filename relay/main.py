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
from typing import Any, List, Optional

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


def _m2o_name(value: Any) -> str:
    """Anzeigename eines many2one-Feldes robust ermitteln ([id, name] / {…} / False)."""
    if isinstance(value, list) and len(value) >= 2:
        return value[1] or ""
    if isinstance(value, dict):
        return value.get("display_name") or value.get("name") or ""
    return ""


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


def _nl2br(text: str) -> str:
    """Text HTML-sicher machen und Zeilenumbrüche als <br/> darstellen (Text wird NIE verändert)."""
    escaped = _html_escape((text or "").replace("\r\n", "\n").replace("\r", "\n"))
    return escaped.replace("\n", "<br/>")


# Marker, an denen der zitierte Verlauf üblicherweise beginnt (DE + EN). Reine Heuristik.
_QUOTE_PATTERNS = [re.compile(p, re.IGNORECASE) for p in [
    r"^-{2,}\s*(Urspr[uü]ngliche Nachricht|Original Message)\s*-{2,}\s*$",
    r"^_{5,}\s*$",
    r"^Am\s.+\sschrieb.+:\s*$",
    r"^On\s.+\swrote:\s*$",
    r"^(Von|From):\s.+$",
    r"^>.*$",
]]


def _extract_last_message(text: str) -> str:
    """Schneidet den Text an der ersten erkannten Zitat-Grenze ab. Es wird nur GESCHNITTEN, nie umgeschrieben."""
    if not text:
        return text
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    cut = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped and any(rx.match(stripped) for rx in _QUOTE_PATTERNS):
            cut = i
            break
    if not cut:  # None oder 0 -> nichts Sinnvolles gefunden, ganzen Text behalten
        return text
    return "\n".join(lines[:cut]).rstrip()


# --- Ganzer Verlauf: Trenner zwischen den einzelnen Nachrichten einfügen ---

# Kopfzeilen eines zitierten Nachrichtenkopfs (Outlook/Gmail, DE + EN)
_HEADER_FIELD = re.compile(
    r"^(Von|Gesendet|An|Betreff|CC|Cc|From|Sent|To|Subject|Date|Datum|Reply-To|Antwort an):",
    re.IGNORECASE,
)

# "Harte" Nachrichtengrenzen (nicht die einzelnen >-Zeilen)
_HARD_BOUNDARY = [re.compile(p, re.IGNORECASE) for p in [
    r"^-{2,}\s*(Urspr[uü]ngliche Nachricht|Original Message)\s*-{2,}\s*$",
    r"^_{5,}\s*$",
    r"^Am\s.+\sschrieb.+:\s*$",
    r"^On\s.+\swrote:\s*$",
]]

_DIVIDER = "──────────── vorherige Nachricht ────────────"


def _format_thread_html(text: str) -> str:
    """Fügt vor jeder erkannten Nachrichtengrenze einen sichtbaren Trenner ein.
    Der eigentliche Text wird dabei nur escaped, nie verändert."""
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: List[str] = []
    armed = True       # bereit, die nächste Grenze zu erkennen
    in_quote = False
    for line in lines:
        stripped = line.strip()
        is_quote = stripped.startswith(">")
        is_hard = any(rx.match(stripped) for rx in _HARD_BOUNDARY)
        is_header = bool(_HEADER_FIELD.match(stripped))

        if armed and out and (is_hard or is_header or (is_quote and not in_quote)):
            out.append("")
            out.append(_DIVIDER)
            out.append("")
            armed = False

        out.append(_html_escape(line))

        # Erst wieder "scharf" schalten, sobald echter Nachrichtentext folgt
        if stripped and not is_quote and not is_hard and not is_header:
            armed = True
        in_quote = is_quote

    return "<br/>".join(out)


def _build_note_html(meta: "NoteMeta", body_html: str, attachments: List[str]) -> str:
    parts = ["<p><b>E-Mail archiviert</b></p>"]
    header = []
    if meta.sender:  header.append("<b>Von:</b> " + _html_escape(meta.sender))
    if meta.to:      header.append("<b>An:</b> " + _html_escape(meta.to))
    if meta.cc:      header.append("<b>CC:</b> " + _html_escape(meta.cc))
    if meta.date:    header.append("<b>Datum:</b> " + _html_escape(meta.date))
    if meta.subject: header.append("<b>Betreff:</b> " + _html_escape(meta.subject))
    if header:
        parts.append("<p>" + "<br/>".join(header) + "</p>")
    names = ", ".join(_html_escape(a) for a in (attachments or []) if a)
    if names:
        parts.append("<p><b>Anhänge:</b> " + names + "</p>")
    parts.append("<hr/>")
    parts.append("<div>" + body_html + "</div>")
    return "".join(parts)


class PartnerSearch(BaseModel):
    query: str


class TargetSearch(BaseModel):
    type: str                        # contact | project | task | todo | sale_order | opportunity
    query: str = ""
    project_id: Optional[int] = None  # nur für type=task


# Modelle, in deren Chatter das Add-in schreiben darf
ALLOWED_MODELS = {"res.partner", "project.task", "sale.order", "crm.lead"}

# Verkaufsauftrags-Status in lesbarer Form
SALE_STATES = {
    "draft": "Angebot",
    "sent": "Angebot gesendet",
    "sale": "Verkaufsauftrag",
    "done": "Gesperrt",
    "cancel": "Storniert",
}


def _resolve_target(res_model: Optional[str], res_id: Optional[int],
                    partner_id: Optional[int]) -> tuple:
    """Zielmodell + Datensatz-ID bestimmen; partner_id bleibt als Altform gültig."""
    if res_model and res_id:
        if res_model not in ALLOWED_MODELS:
            raise HTTPException(status_code=400, detail=f"Modell nicht erlaubt: {res_model}")
        return res_model, res_id
    if partner_id:
        return "res.partner", partner_id
    raise HTTPException(status_code=400, detail="Kein Ziel angegeben (res_model/res_id fehlen).")


def _record_url(model: str, rid: int) -> str:
    return f"{ODOO_BASE_URL}/web#id={rid}&model={model}&view_type=form"


class EmlAttach(BaseModel):
    partner_id: Optional[int] = None   # Altform (Kontakt)
    res_model: Optional[str] = None
    res_id: Optional[int] = None
    filename: str
    eml_base64: str
    subject: str = ""


class NoteMeta(BaseModel):
    subject: str = ""
    sender: str = ""
    to: str = ""
    cc: str = ""
    date: str = ""


class ChatterNote(BaseModel):
    partner_id: Optional[int] = None   # Altform (Kontakt)
    res_model: Optional[str] = None
    res_id: Optional[int] = None
    scope: str = "all"          # "all" = ganzer Verlauf, "last" = nur letzte Nachricht
    body_text: str = ""
    meta: NoteMeta
    attachments: List[str] = []


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


@app.post("/targets/search")
async def targets_search(
    body: TargetSearch,
    x_client_token: Optional[str] = Header(default=None),
):
    """Sucht Ziel-Datensätze je nach Typ: Kontakt, Projekt, Aufgabe, ToDo, Verkaufsauftrag, Verkaufschance."""
    _check_token(x_client_token)

    t = (body.type or "").strip()
    q = (body.query or "").strip()
    results = []

    if t == "contact":
        if len(q) < 2:
            return {"results": []}
        rows = await _odoo_call("res.partner", "search_read", {
            "domain": ["|", ["name", "ilike", q], ["email", "ilike", q]],
            "fields": ["name", "email", "parent_id", "commercial_company_name", "is_company"],
            "limit": 20,
            "order": "name asc",
        })
        for r in (rows or []):
            name = r.get("name") or ""
            if r.get("is_company"):
                name += " (Firma)"
            meta = " · ".join(x for x in [r.get("email") or "", _company_of(r)] if x)
            results.append({"id": r.get("id"), "name": name, "meta": meta})

    elif t == "project":
        domain = [["name", "ilike", q]] if q else []
        rows = await _odoo_call("project.project", "search_read", {
            "domain": domain,
            "fields": ["name", "partner_id"],
            "limit": 40,
            "order": "name asc",
        })
        for r in (rows or []):
            results.append({"id": r.get("id"), "name": r.get("name") or "",
                            "meta": _m2o_name(r.get("partner_id"))})

    elif t == "task":
        if not body.project_id:
            raise HTTPException(status_code=400, detail="project_id fehlt für die Aufgabensuche.")
        domain: List[Any] = [["project_id", "=", body.project_id]]
        if q:
            domain.append(["name", "ilike", q])
        rows = await _odoo_call("project.task", "search_read", {
            "domain": domain,
            "fields": ["name", "stage_id"],
            "limit": 40,
            "order": "name asc",
        })
        for r in (rows or []):
            results.append({"id": r.get("id"), "name": r.get("name") or "",
                            "meta": _m2o_name(r.get("stage_id"))})

    elif t == "todo":
        # ToDos sind in Odoo Aufgaben ohne Projekt
        domain = [["project_id", "=", False]]
        if q:
            domain.append(["name", "ilike", q])
        rows = await _odoo_call("project.task", "search_read", {
            "domain": domain,
            "fields": ["name", "date_deadline"],
            "limit": 40,
            "order": "id desc",
        })
        for r in (rows or []):
            deadline = r.get("date_deadline") or ""
            results.append({"id": r.get("id"), "name": r.get("name") or "",
                            "meta": ("Frist: " + str(deadline)) if deadline else ""})

    elif t == "sale_order":
        if len(q) < 2:
            return {"results": []}
        rows = await _odoo_call("sale.order", "search_read", {
            "domain": ["|", ["name", "ilike", q], ["partner_id", "ilike", q]],
            "fields": ["name", "partner_id", "state"],
            "limit": 20,
            "order": "id desc",
        })
        for r in (rows or []):
            state = SALE_STATES.get(r.get("state") or "", r.get("state") or "")
            meta = " · ".join(x for x in [_m2o_name(r.get("partner_id")), state] if x)
            results.append({"id": r.get("id"), "name": r.get("name") or "", "meta": meta})

    elif t == "opportunity":
        if len(q) < 2:
            return {"results": []}
        rows = await _odoo_call("crm.lead", "search_read", {
            "domain": ["&", ["type", "=", "opportunity"],
                       "|", "|",
                       ["name", "ilike", q],
                       ["partner_id", "ilike", q],
                       ["partner_name", "ilike", q]],
            "fields": ["name", "partner_id", "partner_name", "stage_id"],
            "limit": 20,
            "order": "id desc",
        })
        for r in (rows or []):
            client = _m2o_name(r.get("partner_id")) or (r.get("partner_name") or "")
            meta = " · ".join(x for x in [client, _m2o_name(r.get("stage_id"))] if x)
            results.append({"id": r.get("id"), "name": r.get("name") or "", "meta": meta})

    else:
        raise HTTPException(status_code=400, detail=f"Unbekannter Suchtyp: {t}")

    return {"results": results}


@app.post("/chatter/eml")
async def chatter_eml(
    body: EmlAttach,
    x_client_token: Optional[str] = Header(default=None),
):
    """Hängt die Original-E-Mail als .eml-Datei an die Chatter des Ziel-Datensatzes an."""
    _check_token(x_client_token)

    model, rid = _resolve_target(body.res_model, body.res_id, body.partner_id)
    filename = _safe_filename(body.filename)

    # 1) Anhang direkt am Ziel-Datensatz anlegen
    attachment_vals = {
        "name": filename,
        "datas": body.eml_base64,
        "mimetype": "message/rfc822",
        "res_model": model,
        "res_id": rid,
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
    post_result = await _odoo_call(model, "message_post", {
        "ids": [rid],
        "body": note_body,
        "body_is_html": True,  # sonst behandelt Odoo den String als Text und zeigt HTML-Tags wörtlich
        "message_type": "comment",
        "subtype_xmlid": "mail.mt_note",
        "attachment_ids": [attachment_id],
    })

    return {
        "ok": True,
        "attachment_id": attachment_id,
        "message_id": _extract_id(post_result),
        "partner_url": _record_url(model, rid),
    }


@app.post("/chatter/note")
async def chatter_note(
    body: ChatterNote,
    x_client_token: Optional[str] = Header(default=None),
):
    """Postet eine saubere Text-Notiz (ohne KI) in die Chatter des Ziel-Datensatzes."""
    _check_token(x_client_token)

    model, rid = _resolve_target(body.res_model, body.res_id, body.partner_id)

    text = body.body_text or ""
    if body.scope == "last":
        body_html = _nl2br(_extract_last_message(text))
    else:
        body_html = _format_thread_html(text)

    note_html = _build_note_html(body.meta, body_html, body.attachments)
    post_result = await _odoo_call(model, "message_post", {
        "ids": [rid],
        "body": note_html,
        "body_is_html": True,
        "message_type": "comment",
        "subtype_xmlid": "mail.mt_note",
    })

    return {
        "ok": True,
        "message_id": _extract_id(post_result),
        "partner_url": _record_url(model, rid),
    }
