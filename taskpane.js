/* "An Odoo senden" – Taskpane
 *
 * M3 + Sicherheits-Gate: Kontaktsuche über das Railway-Relay -> Odoo.
 * Der Zugriffs-Token steht NICHT mehr im Code, sondern wird einmalig
 * eingegeben und in den roamingSettings des Postfachs gespeichert.
 * Der Odoo-API-Key liegt weiterhin ausschließlich auf Railway.
 */

/* ---------- Konfiguration ---------- */
var RELAY_BASE_URL = "https://aluthermo.up.railway.app";  // nicht geheim
var TOKEN_KEY = "clientToken";

var clientToken = "";
var selectedPartner = null;
var searchTimer = null;

Office.onReady(function (info) {
  if (!(info && info.host === Office.HostType.Outlook)) {
    setStatus("ok", "Vorschau im Browser ✓ – in Outlook öffnen, um eine E-Mail zu laden.");
    return;
  }
  setStatus("ok", "Add-in bereit ✓");

  clientToken = Office.context.roamingSettings.get(TOKEN_KEY) || "";

  setupSettingsUi();
  setupChoiceUi();
  setupContactSearch();
  loadItemDetails();

  if (clientToken) {
    showMainFlow();
    autoSearchSender();
  } else {
    setText("settings-hint", "Bitte einmalig den Zugriffs-Token eingeben, um Kontakte zu suchen.");
    showSettings();
  }
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

function metaOf(p) { return [p.email, p.company].filter(Boolean).join(" · ") || "—"; }
function nameOf(p) { return p.name + (p.is_company ? " (Firma)" : ""); }

/* ---------- Einstellungen / Token ---------- */

function setupSettingsUi() {
  document.getElementById("btn-settings").addEventListener("click", function () {
    showSettings();
  });

  document.getElementById("settings-show").addEventListener("change", function (e) {
    document.getElementById("settings-token").type = e.target.checked ? "text" : "password";
  });

  document.getElementById("settings-save").addEventListener("click", saveToken);

  document.getElementById("settings-cancel").addEventListener("click", function () {
    if (clientToken) {
      showMainFlow();
    } else {
      setText("settings-status", "Ohne Token kann das Add-in keine Kontakte suchen.");
    }
  });
}

function showSettings() {
  document.getElementById("settings").style.display = "block";
  document.getElementById("main-flow").style.display = "none";
  document.getElementById("settings-token").value = clientToken || "";
  setText("settings-status", "");
}

function showMainFlow() {
  document.getElementById("settings").style.display = "none";
  document.getElementById("main-flow").style.display = "block";
}

function saveToken() {
  var value = (document.getElementById("settings-token").value || "").trim();
  if (!value) {
    setText("settings-status", "Bitte einen Token eingeben.");
    return;
  }
  Office.context.roamingSettings.set(TOKEN_KEY, value);
  Office.context.roamingSettings.saveAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      clientToken = value;
      setText("settings-status", "");
      showMainFlow();
      autoSearchSender();
    } else {
      setText("settings-status", "Speichern fehlgeschlagen: " +
        (res.error && res.error.message ? res.error.message : "unbekannter Fehler"));
    }
  });
}

/* ---------- E-Mail einlesen ---------- */

function loadItemDetails() {
  var item = Office.context.mailbox.item;

  setText("f-subject", item.subject || "(kein Betreff)");
  setText("f-from", formatAddress(item.from) || "—");
  setText("f-to", formatList(item.to));
  setText("f-cc", formatList(item.cc));

  var d = item.dateTimeCreated;
  setText("f-date", d ? new Date(d).toLocaleString("de-DE") : "—");

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

function autoSearchSender() {
  var item = Office.context.mailbox.item;
  var fromEmail = item.from && item.from.emailAddress;
  if (fromEmail) {
    document.getElementById("contact-search").value = fromEmail;
    doSearch(fromEmail);
  }
}

/* ---------- Auswahl-Oberfläche (Text/.eml, Umfang) ---------- */

function setupChoiceUi() {
  var textRadio = document.getElementById("mode-text");
  var emlRadio = document.getElementById("mode-eml");
  var emlLabel = document.getElementById("mode-eml-label");
  var scopeBox = document.getElementById("scope-options");
  var scopeLast = document.getElementById("scope-last");
  var scopeAll = document.getElementById("scope-all");

  var emlSupported = Office.context.requirements.isSetSupported("Mailbox", "1.14");
  if (!emlSupported) {
    emlRadio.disabled = true;
    emlLabel.classList.add("is-disabled");
    emlLabel.title = "Dieser Outlook-Client ist zu alt für den .eml-Export (benötigt Mailbox 1.14).";
    emlLabel.appendChild(document.createTextNode("  (in diesem Client nicht verfügbar)"));
  }

  function refresh() {
    scopeBox.style.display = textRadio.checked ? "block" : "none";
    updateSummary();
  }

  [textRadio, emlRadio, scopeLast, scopeAll].forEach(function (r) {
    if (r) { r.addEventListener("change", refresh); }
  });

  document.getElementById("btn-cancel").addEventListener("click", function () {
    textRadio.checked = true;
    scopeLast.checked = true;
    refresh();
  });

  refresh();
}

function updateSummary() {
  var textMode = document.getElementById("mode-text").checked;
  var parts = ["Auswahl: " + (textMode ? "Text in Chatter" : "E-Mail (.eml) anhängen")];
  if (textMode) {
    parts.push(document.getElementById("scope-last").checked ? "Nur letzte Nachricht" : "Ganzer Verlauf");
  }
  parts.push(selectedPartner ? ("Kontakt: " + selectedPartner.name) : "Kein Kontakt gewählt");
  parts.push("Senden folgt (M4/M5)");
  setText("selection-summary", parts.join(" · "));
}

/* ---------- Kontaktsuche (Relay -> Odoo) ---------- */

function setupContactSearch() {
  var input = document.getElementById("contact-search");

  input.addEventListener("input", function () {
    clearTimeout(searchTimer);
    var q = input.value.trim();
    if (q.length < 2) { renderResults([]); setResultsStatus(""); return; }
    setResultsStatus("Suche…");
    searchTimer = setTimeout(function () { doSearch(q); }, 350);
  });

  document.getElementById("sel-change").addEventListener("click", function () {
    selectedPartner = null;
    document.getElementById("contact-selected").style.display = "none";
    document.getElementById("contact-picker").style.display = "block";
    updateSummary();
  });
}

function setResultsStatus(msg) { setText("contact-status", msg); }

function doSearch(query) {
  if (!clientToken) {
    setResultsStatus("Bitte zuerst den Zugriffs-Token in den Einstellungen eingeben.");
    showSettings();
    return;
  }
  fetch(RELAY_BASE_URL + "/partners/search", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Client-Token": clientToken
    },
    body: JSON.stringify({ query: query })
  }).then(function (res) {
    if (res.status === 401) { throw new Error("401"); }
    if (!res.ok) {
      return res.text().then(function (t) {
        throw new Error("Relay " + res.status + ": " + t.slice(0, 200));
      });
    }
    return res.json();
  }).then(function (data) {
    renderResults((data && data.partners) || []);
  }).catch(function (err) {
    renderResults([]);
    if (err.message === "401") {
      setResultsStatus("Zugriffs-Token ungültig – bitte in den Einstellungen prüfen.");
      showSettings();
    } else {
      setResultsStatus("Fehler bei der Suche: " + err.message);
    }
  });
}

function renderResults(list) {
  var ul = document.getElementById("contact-results");
  ul.innerHTML = "";
  if (!list.length) {
    if (!document.getElementById("contact-status").textContent) {
      setResultsStatus("Keine Treffer");
    }
    return;
  }
  setResultsStatus(list.length + " Treffer");
  list.forEach(function (p) {
    var li = document.createElement("li");
    li.className = "results__item";

    var nameDiv = document.createElement("div");
    nameDiv.className = "results__name";
    nameDiv.textContent = nameOf(p);

    var metaDiv = document.createElement("div");
    metaDiv.className = "results__meta";
    metaDiv.textContent = metaOf(p);

    li.appendChild(nameDiv);
    li.appendChild(metaDiv);
    li.addEventListener("click", function () { selectPartner(p); });
    ul.appendChild(li);
  });
}

function selectPartner(p) {
  selectedPartner = p;
  document.getElementById("contact-picker").style.display = "none";
  document.getElementById("contact-selected").style.display = "flex";
  setText("sel-name", nameOf(p));
  setText("sel-meta", metaOf(p));
  updateSummary();
}
