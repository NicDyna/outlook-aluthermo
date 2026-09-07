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

import base64
import binascii
import json
import logging
import os
import re
import secrets
import time
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

# Fehlerdetails landen im Railway-Log (Deployments -> View Logs), nie beim Aufrufer
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("relay")

# --- Größen-Grenzen (greifen, BEVOR der Body gelesen wird) ---

# Größte E-Mail, die als .eml archiviert werden darf (entpackt, nicht base64).
# Hinweis: eine Anfrage an dieser Grenze belegt auf Railway kurzzeitig ein
# Mehrfaches davon an Arbeitsspeicher. Startet das Relay bei großen Anhängen
# neu, diesen Wert verkleinern; 25 MB decken praktisch jede Outlook-Mail ab.
MAX_EML_BYTES = 50 * 1024 * 1024

# base64 ist rund 4/3 so groß wie die Rohdaten, plus etwas Reserve.
MAX_EML_B64_CHARS = (MAX_EML_BYTES // 3 + 1) * 4 + 4096

# Obergrenze für den gesamten HTTP-Body, je Endpunkt.
MAX_BODY_BYTES = {
    "/chatter/eml": MAX_EML_B64_CHARS + 64 * 1024,   # Anhang + Kopfdaten
    "/chatter/note": 2 * 1024 * 1024,                # nur Text
}
DEFAULT_MAX_BODY_BYTES = 64 * 1024                   # Suche, Benutzerliste

# Einziger Endpunkt, der ohne Token erreichbar ist.
OPEN_PATHS = {"/health"}


def _token_ok(token: Optional[str]) -> bool:
    """Zeitkonstanter Vergleich; ohne gesetzten CLIENT_TOKEN wird alles abgelehnt."""
    if not CLIENT_TOKEN or not token:
        return False
    return secrets.compare_digest(token.encode("utf-8"), CLIENT_TOKEN.encode("utf-8"))


def _size_message(limit: int) -> str:
    unit = "MB" if limit >= 1024 * 1024 else "KB"
    value = limit // (1024 * 1024) if unit == "MB" else limit // 1024
    return f"Anfrage zu groß (Grenze: {value} {unit})."


async def _send_json(send, status: int, detail: str) -> None:
    """Antwort direkt auf ASGI-Ebene senden, ohne den Body gelesen zu haben."""
    payload = json.dumps({"detail": detail}).encode("utf-8")
    await send({
        "type": "http.response.start",
        "status": status,
        "headers": [
            (b"content-type", b"application/json; charset=utf-8"),
            (b"content-length", str(len(payload)).encode("ascii")),
        ],
    })
    await send({"type": "http.response.body", "body": payload})


class TokenAndSizeGate:
    """Prüft Token und Größe, BEVOR FastAPI den Request-Body einliest.

    Der Token-Test in den Endpunkten selbst kommt zu spät: FastAPI liest und
    prüft den kompletten Body, bevor die Funktion überhaupt startet. Ohne diese
    Middleware könnte also jeder ohne Token beliebig große Daten schicken und
    das Relay lahmlegen. Hier fällt die Entscheidung, bevor ein einziges Byte
    des Bodys gelesen wird; uvicorn stoppt dann von selbst bei rund 64 KB.
    """

    def __init__(self, app):
        self.app = app

    @staticmethod
    def _header(scope, name: bytes) -> Optional[str]:
        """Header aus dem rohen ASGI-Scope lesen (Namen sind dort klein geschrieben)."""
        for key, value in scope.get("headers", []):
            if key == name:
                return value.decode("latin-1", "replace")
        return None

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        path = scope.get("path", "")
        # CORS-Vorabfragen tragen nie einen Token und müssen durchgelassen werden.
        if scope.get("method") == "OPTIONS" or path in OPEN_PATHS:
            await self.app(scope, receive, send)
            return

        if not _token_ok(self._header(scope, b"x-client-token")):
            await _send_json(send, 401, "Ungültiger oder fehlender Client-Token.")
            return

        limit = MAX_BODY_BYTES.get(path, DEFAULT_MAX_BODY_BYTES)

        declared = self._header(scope, b"content-length")
        if declared is not None:
            try:
                length = int(declared)
            except ValueError:
                await _send_json(send, 400, "Ungültige Content-Length-Angabe.")
                return
            if length > limit:
                await _send_json(send, 413, _size_message(limit))
                return
            await self.app(scope, receive, send)
            return

        # Ohne Content-Length (chunked): selbst mitzählen und notfalls abbrechen,
        # bevor die Daten an FastAPI weitergereicht werden.
        chunks: List[bytes] = []
        total = 0
        while True:
            message = await receive()
            if message["type"] != "http.request":
                break
            chunk = message.get("body", b"")
            total += len(chunk)
            if total > limit:
                await _send_json(send, 413, _size_message(limit))
                return
            chunks.append(chunk)
            if not message.get("more_body", False):
                break

        buffered = b"".join(chunks)
        replayed = False

        async def replay():
            nonlocal replayed
            if replayed:
                return {"type": "http.disconnect"}
            replayed = True
            return {"type": "http.request", "body": buffered, "more_body": False}

        await self.app(scope, replay, send)


app = FastAPI(title="Outlook -> Odoo Relay", version="0.1.0")

# Reihenfolge beachten: zuletzt hinzugefügt liegt außen. CORS muss außen liegen,
# damit auch die 401/413-Antworten der Gate-Middleware die CORS-Header bekommen.
app.add_middleware(TokenAndSizeGate)

# Nur die GitHub-Pages-Herkunft darf das Relay aus dem Browser aufrufen.
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "X-Client-Token"],
)


def _check_token(token: Optional[str]) -> None:
    """Zweite Absicherung; die Gate-Middleware hat den Token bereits geprüft."""
    if not _token_ok(token):
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
        # Details nur ins Log, nicht an den Aufrufer (keine Odoo-Interna preisgeben)
        log.error("Odoo-Fehler %s bei %s/%s: %s", resp.status_code, model, method, resp.text[:1000])
        raise HTTPException(
            status_code=502,
            detail=f"Odoo-Fehler ({resp.status_code}) – Details stehen im Railway-Log.",
        )
    return resp.json()


def _m2o_name(value: Any) -> str:
    """Anzeigename eines many2one-Feldes robust ermitteln ([id, name] / {…} / False)."""
    if isinstance(value, list) and len(value) >= 2:
        return value[1] or ""
    if isinstance(value, dict):
        return value.get("display_name") or value.get("name") or ""
    return ""


def _m2o_id(value: Any) -> Optional[int]:
    """ID eines many2one-Feldes robust ermitteln ([id, name] / {…} / int / False)."""
    if isinstance(value, bool):      # Odoo liefert False für leere many2one-Felder
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, list) and value and isinstance(value[0], int):
        return value[0]
    if isinstance(value, dict) and isinstance(value.get("id"), int):
        return value["id"]
    return None


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


def _escape_like(value: str) -> str:
    """% und _ entwerten, damit eine Suche nach "%" nicht den ganzen Bestand findet.

    Odoo baut daraus ein SQL-ILIKE ohne eigene ESCAPE-Angabe; dort ist der
    Backslash das Standard-Escape-Zeichen (odoo/orm/fields.py, Branch 19.0).
    """
    return re.sub(r"([\\%_])", r"\\\1", value)


def _eml_too_large() -> str:
    return f"E-Mail zu groß (max. {MAX_EML_BYTES // (1024 * 1024)} MB)."


# Eine .eml beginnt immer mit einer Kopfzeile, etwa "Received:" oder "From:".
_EML_FIRST_HEADER = re.compile(rb"^[A-Za-z][A-Za-z0-9-]{0,60}:")


def _decode_eml(data: str) -> str:
    """Anhang prüfen, bevor etwas an Odoo geht: Größe, gültiges base64, E-Mail-Form.

    Gibt das normalisierte base64 zurück, damit Odoo genau die Daten bekommt,
    die hier geprüft wurden.
    """
    if len(data or "") > MAX_EML_B64_CHARS:
        raise HTTPException(status_code=413, detail=_eml_too_large())
    try:
        raw = base64.b64decode(data or "", validate=True)
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="Anhang ist kein gültiges base64.")
    if not raw:
        raise HTTPException(status_code=400, detail="Anhang ist leer.")
    if len(raw) > MAX_EML_BYTES:
        raise HTTPException(status_code=413, detail=_eml_too_large())
    if not _EML_FIRST_HEADER.match(raw[:1024]):
        raise HTTPException(
            status_code=400,
            detail="Anhang sieht nicht wie eine E-Mail aus (keine Kopfzeile am Anfang).",
        )
    return base64.b64encode(raw).decode("ascii")


def _html_escape(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _nl2br(text: str) -> str:
    """Text HTML-sicher machen und Zeilenumbrüche als <br/> darstellen (Text wird NIE verändert)."""
    escaped = _html_escape((text or "").replace("\r\n", "\n").replace("\r", "\n"))
    return escaped.replace("\n", "<br/>")


# Zitat-Kopfzeilen sind immer kurz. Längere Zeilen werden gar nicht erst geprüft,
# damit eine einzelne sehr lange Zeile die Regex-Suche nicht ausbremsen kann.
MAX_MATCH_LINE = 500

# Marker, an denen der zitierte Verlauf üblicherweise beginnt (DE + EN). Reine Heuristik.
# Die Platzhalter sind bewusst begrenzt (.{0,200} statt .+): zwei unbegrenzte
# Platzhalter in einer Zeile lassen die Laufzeit quadratisch wachsen.
_QUOTE_PATTERNS = [re.compile(p, re.IGNORECASE) for p in [
    r"^-{2,}\s*(Urspr[uü]ngliche Nachricht|Original Message)\s*-{2,}\s*$",
    r"^_{5,}\s*$",
    r"^Am\s.{1,200}?\sschrieb.{0,200}:\s*$",
    r"^On\s.{1,200}?\swrote:\s*$",
    r"^(Von|From):\s.+$",
    r"^>.*$",
]]


def _matches_any(stripped: str, patterns) -> bool:
    """Zeile gegen die Marker prüfen; eine zu lange Zeile ist nie ein Marker."""
    if not stripped or len(stripped) > MAX_MATCH_LINE:
        return False
    return any(rx.match(stripped) for rx in patterns)


def _extract_last_message(text: str) -> str:
    """Schneidet den Text an der ersten erkannten Zitat-Grenze ab. Es wird nur GESCHNITTEN, nie umgeschrieben."""
    if not text:
        return text
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    cut = None
    for i, line in enumerate(lines):
        if _matches_any(line.strip(), _QUOTE_PATTERNS):
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

# "Harte" Nachrichtengrenzen (nicht die einzelnen >-Zeilen), ebenfalls begrenzt
_HARD_BOUNDARY = [re.compile(p, re.IGNORECASE) for p in [
    r"^-{2,}\s*(Urspr[uü]ngliche Nachricht|Original Message)\s*-{2,}\s*$",
    r"^_{5,}\s*$",
    r"^Am\s.{1,200}?\sschrieb.{0,200}:\s*$",
    r"^On\s.{1,200}?\swrote:\s*$",
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
        is_hard = _matches_any(stripped, _HARD_BOUNDARY)
        is_header = _matches_any(stripped, (_HEADER_FIELD,))

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


class UsersRequest(BaseModel):
    # E-Mail-Adresse des angemeldeten Outlook-Postfachs, nur für den Vorschlag
    mailbox: str = ""


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


# --- Absender: unter welchem Odoo-Benutzer die Notiz erscheinen darf ---

# Die Benutzerliste ändert sich selten; kurzes Zwischenspeichern spart bei jedem
# Senden eine zusätzliche Odoo-Abfrage.
USER_CACHE_TTL = 300  # Sekunden
_user_cache: dict = {"at": 0.0, "users": []}


async def _internal_users(force: bool = False) -> List[dict]:
    """Aktive interne Odoo-Benutzer laden (ohne Portal- und Public-Benutzer)."""
    now = time.monotonic()
    if not force and _user_cache["users"] and (now - _user_cache["at"]) < USER_CACHE_TTL:
        return _user_cache["users"]

    rows = await _odoo_call("res.users", "search_read", {
        "domain": [["share", "=", False], ["active", "=", True]],
        "fields": ["name", "login", "email", "partner_id"],
        "limit": 200,
        "order": "name asc",
    })

    users = []
    for r in (rows or []):
        partner_id = _m2o_id(r.get("partner_id"))
        if not partner_id:
            continue
        users.append({
            "id": r.get("id"),
            "partner_id": partner_id,
            "name": r.get("name") or "",
            "login": r.get("login") or "",
            "email": r.get("email") or "",
        })

    _user_cache["users"] = users
    _user_cache["at"] = now
    return users


async def _resolve_author(author_id: Optional[int]) -> Optional[int]:
    """Prüft, dass der gewünschte Absender ein interner Odoo-Benutzer ist.

    Ohne diese Prüfung könnte mit dem Client-Token eine Notiz im Namen eines
    beliebigen Partners erscheinen, zum Beispiel im Namen eines Kunden.
    Ohne author_id bleibt alles wie bisher: der technische API-Benutzer
    ist dann der Autor.
    """
    if not author_id:
        return None
    if author_id not in {u["partner_id"] for u in await _internal_users()}:
        # Cache könnte veraltet sein (neuer Kollege) -> einmal frisch nachladen
        if author_id not in {u["partner_id"] for u in await _internal_users(force=True)}:
            raise HTTPException(
                status_code=400,
                detail="Ungültiger Absender – bitte erneut auswählen.",
            )
    return author_id


def _post_payload(rid: int, body_html: str, author_id: Optional[int],
                  attachment_ids: Optional[List[int]] = None) -> dict:
    """Einheitliche Argumente für message_post (interne Notiz, kein Mailversand)."""
    payload: dict = {
        "ids": [rid],
        "body": body_html,
        # sonst behandelt Odoo den String als Text und zeigt HTML-Tags wörtlich
        "body_is_html": True,
        "message_type": "comment",
        "subtype_xmlid": "mail.mt_note",
    }
    if attachment_ids:
        payload["attachment_ids"] = attachment_ids
    if author_id:
        # Odoo übernimmt author_id unverändert (mail.thread._message_compute_author),
        # die Notiz erscheint dadurch unter dem gewählten Benutzer.
        payload["author_id"] = author_id
    return payload


class EmlAttach(BaseModel):
    partner_id: Optional[int] = None   # Altform (Kontakt)
    res_model: Optional[str] = None
    res_id: Optional[int] = None
    author_id: Optional[int] = None    # Partner-ID des gewählten Odoo-Benutzers
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
    author_id: Optional[int] = None    # Partner-ID des gewählten Odoo-Benutzers
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

    query = (body.query or "").strip()[:100]
    if len(query) < 2:
        return {"partners": []}

    like = _escape_like(query)
    payload = {
        "domain": ["|", ["name", "ilike", like], ["email", "ilike", like]],
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
    q = (body.query or "").strip()[:100]
    # Platzhalter entwerten: eine Suche nach "%" darf nicht alles zurückgeben.
    like = _escape_like(q)
    results = []

    if t == "contact":
        if len(q) < 2:
            return {"results": []}
        rows = await _odoo_call("res.partner", "search_read", {
            "domain": ["|", ["name", "ilike", like], ["email", "ilike", like]],
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
        domain = [["name", "ilike", like]] if q else []
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
            domain.append(["name", "ilike", like])
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
            domain.append(["name", "ilike", like])
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
            "domain": ["|", ["name", "ilike", like], ["partner_id", "ilike", like]],
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
                       ["name", "ilike", like],
                       ["partner_id", "ilike", like],
                       ["partner_name", "ilike", like]],
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


@app.post("/users/list")
async def users_list(
    body: UsersRequest,
    x_client_token: Optional[str] = Header(default=None),
):
    """Liefert die internen Odoo-Benutzer für die Absender-Auswahl im Add-in.

    Login und E-Mail bleiben im Relay: das Add-in bekommt nur Name und
    partner_id. Der Abgleich mit dem angemeldeten Outlook-Postfach passiert
    hier, das Add-in erhält lediglich den fertigen Vorschlag.
    """
    _check_token(x_client_token)

    users = await _internal_users()
    mailbox = (body.mailbox or "").strip().lower()[:200]
    suggested = None
    if mailbox:
        for u in users:
            if u["login"].lower() == mailbox or u["email"].lower() == mailbox:
                suggested = u["partner_id"]
                break

    return {
        "users": [{"partner_id": u["partner_id"], "name": u["name"]} for u in users],
        "suggested_partner_id": suggested,
    }


@app.post("/chatter/eml")
async def chatter_eml(
    body: EmlAttach,
    x_client_token: Optional[str] = Header(default=None),
):
    """Hängt die Original-E-Mail als .eml-Datei an die Chatter des Ziel-Datensatzes an."""
    _check_token(x_client_token)

    model, rid = _resolve_target(body.res_model, body.res_id, body.partner_id)
    # Absender vor dem Anlegen des Anhangs prüfen, damit bei einem ungültigen
    # Absender kein verwaister Anhang in Odoo zurückbleibt.
    author_id = await _resolve_author(body.author_id)
    filename = _safe_filename(body.filename)
    eml_base64 = _decode_eml(body.eml_base64)

    # 1) Anhang direkt am Ziel-Datensatz anlegen
    attachment_vals = {
        "name": filename,
        "datas": eml_base64,
        "mimetype": "message/rfc822",
        "res_model": model,
        "res_id": rid,
    }
    create_result = await _odoo_call("ir.attachment", "create", {"vals_list": [attachment_vals]})
    attachment_id = _extract_id(create_result)
    if not attachment_id:
        log.error("Anhang-ID nicht erkannt. Odoo-Antwort auf create: %r", create_result)
        raise HTTPException(
            status_code=502,
            detail="Anhang-ID nicht erkannt – Details stehen im Railway-Log.",
        )

    # 2) Interne Chatter-Notiz mit verknüpftem Anhang
    subject = body.subject or filename
    note_body = f"<p>E-Mail archiviert: {_html_escape(subject)}</p>"
    post_result = await _odoo_call(
        model, "message_post",
        _post_payload(rid, note_body, author_id, attachment_ids=[attachment_id]),
    )

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
    author_id = await _resolve_author(body.author_id)

    text = body.body_text or ""
    if body.scope == "last":
        body_html = _nl2br(_extract_last_message(text))
    else:
        body_html = _format_thread_html(text)

    note_html = _build_note_html(body.meta, body_html, body.attachments)
    post_result = await _odoo_call(
        model, "message_post",
        _post_payload(rid, note_html, author_id),
    )

    return {
        "ok": True,
        "message_id": _extract_id(post_result),
        "partner_url": _record_url(model, rid),
    }
