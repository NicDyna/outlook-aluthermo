/* "An Odoo senden" – Taskpane
 *
 * M2: zeigt die echten Details der geöffneten E-Mail (Betreff, Von, An, CC,
 * Datum, Textvorschau) und legt die Auswahl-Oberfläche an (Text / .eml,
 * Umfang, Kontakt-Platzhalter). Noch KEINE Odoo-Anbindung.
 */
Office.onReady(function (info) {
  if (!(info && info.host === Office.HostType.Outlook)) {
    // Direkt im Browser geöffnet – keine E-Mail vorhanden
    setStatus("ok", "Vorschau im Browser ✓ – in Outlook öffnen, um eine E-Mail zu laden.");
    return;
  }

  setStatus("ok", "Add-in bereit ✓");
  loadItem();
  setupChoiceUi();
});

/* ---------- Helfer ---------- */

function setStatus(state, message) {
  var statusEl = document.getElementById("status");
  var textEl = document.getElementById("status-text");
  statusEl.classList.remove("status--loading", "status--ok", "status--error");
  statusEl.classList.add("status--" + state);
  textEl.textContent = message;
}

function setText(id, value) {
  var el = document.getElementById(id);
  if (el) { el.textContent = value; }
}

function formatAddress(a) {
  if (!a) { return ""; }
  var name = a.displayName || "";
  var email = a.emailAddress || "";
  if (name && email && name !== email) { return name + " <" + email + ">"; }
  return email || name || "";
}

function formatList(arr) {
  if (!arr || !arr.length) { return "—"; }
  return arr.map(formatAddress).join(", ");
}

/* ---------- E-Mail einlesen ---------- */

function loadItem() {
  var item = Office.context.mailbox.item;

  setText("f-subject", item.subject || "(kein Betreff)");
  setText("f-from", formatAddress(item.from) || "—");
  setText("f-to", formatList(item.to));
  setText("f-cc", formatList(item.cc));

  var d = item.dateTimeCreated;
  setText("f-date", d ? new Date(d).toLocaleString("de-DE") : "—");

  // Reiner Text des Nachrichtentexts (asynchron)
  item.body.getAsync(Office.CoercionType.Text, function (res) {
    var el = document.getElementById("f-body");
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      el.textContent = res.value && res.value.trim() ? res.value : "(leerer Text)";
    } else {
      el.textContent = "Text konnte nicht gelesen werden: " +
        (res.error && res.error.message ? res.error.message : "unbekannter Fehler");
    }
  });
}

/* ---------- Auswahl-Oberfläche ---------- */

function setupChoiceUi() {
  var textRadio = document.getElementById("mode-text");
  var emlRadio = document.getElementById("mode-eml");
  var emlLabel = document.getElementById("mode-eml-label");
  var scopeBox = document.getElementById("scope-options");
  var scopeLast = document.getElementById("scope-last");

  // .eml-Export benötigt Mailbox 1.14 – zur Laufzeit prüfen und ggf. sperren
  var emlSupported = Office.context.requirements.isSetSupported("Mailbox", "1.14");
  if (!emlSupported) {
    emlRadio.disabled = true;
    emlLabel.classList.add("is-disabled");
    emlLabel.title = "Dieser Outlook-Client ist zu alt für den .eml-Export (benötigt Mailbox 1.14).";
    emlLabel.appendChild(document.createTextNode("  (in diesem Client nicht verfügbar)"));
  }

  function refresh() {
    // Umfang-Optionen nur bei Text-Modus zeigen
    scopeBox.style.display = textRadio.checked ? "block" : "none";

    // Live-Zusammenfassung der Auswahl
    var summary = "Auswahl: " + (textRadio.checked ? "Text in Chatter" : "E-Mail (.eml) anhängen");
    if (textRadio.checked) {
      summary += " · " + (scopeLast.checked ? "Nur letzte Nachricht" : "Ganzer Verlauf");
    }
    summary += " · Senden wird in den nächsten Schritten aktiviert.";
    setText("selection-summary", summary);
  }

  // Auf alle Auswahländerungen reagieren
  [textRadio, emlRadio, scopeLast, document.getElementById("scope-all")]
    .forEach(function (r) { if (r) { r.addEventListener("change", refresh); } });

  // "Zurücksetzen"
  document.getElementById("btn-cancel").addEventListener("click", function () {
    textRadio.checked = true;
    scopeLast.checked = true;
    refresh();
  });

  refresh();
}
