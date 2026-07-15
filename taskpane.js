/* "An Odoo senden" – Taskpane
 *
 * M4 + M5 (ohne KI): .eml anhängen ODER saubere Text-Notiz in die Chatter.
 * Text-Notiz = Kopf (aus Outlook-Daten) + Nachrichtentext; optional nur die
 * letzte Nachricht (regelbasierter Schnitt im Relay, keine KI).
 * Der Text wird dabei nur geschnitten/escaped, niemals umgeschrieben.
 *
 * M6: Ziel wählbar – Kontakt, Projekt→Aufgabe, Verkaufsauftrag, ToDo
 * oder Verkaufschance. Bei Aufgaben erst Projekt, dann Aufgabe wählen.
 */

/* ---------- Konfiguration ---------- */
var RELAY_BASE_URL = "https://aluthermo.up.railway.app";  // nicht geheim
var TOKEN_KEY = "clientToken";

var clientToken = "";
var targetType = "contact";      // contact | task | sale_order | todo | opportunity
var selectedProject = null;      // nur bei targetType === "task"
var selectedTarget = null;       // der Datensatz, in dessen Chatter geschrieben wird
var searchTimer = null;
var projectSearchTimer = null;
var emlSupported = false;
var bodyText = "";

/* Eigenschaften je Zieltyp: Odoo-Modell, Beschriftung, Suchfeld-Platzhalter,
 * minQuery = ab wie vielen Zeichen gesucht wird (0 = Liste erscheint sofort). */
var TARGET_TYPES = {
  contact:     { model: "res.partner",  label: "Kontakt",
                 placeholder: "Kontakt suchen (Name oder E-Mail)…", minQuery: 2 },
  task:        { model: "project.task", label: "Aufgabe",
                 placeholder: "Aufgabe suchen…", minQuery: 0 },
  sale_order:  { model: "sale.order",   label: "Verkaufsauftrag",
                 placeholder: "Auftrag suchen (Nummer oder Kunde)…", minQuery: 2 },
  todo:        { model: "project.task", label: "ToDo",
                 placeholder: "ToDo suchen…", minQuery: 0 },
  opportunity: { model: "crm.lead",     label: "Verkaufschance",
                 placeholder: "Chance suchen (Name oder Kunde)…", minQuery: 2 }
};

Office.onReady(function (info) {
  if (!(info && info.host === Office.HostType.Outlook)) {
    setStatus("ok", "Vorschau im Browser ✓ – in Outlook öffnen, um eine E-Mail zu laden.");
    return;
  }
  setStatus("ok", "Add-in bereit ✓");

  clientToken = Office.context.roamingSettings.get(TOKEN_KEY) || "";

  setupSettingsUi();
  setupChoiceUi();
  setupTargetUi();
  loadItemDetails();

  if (clientToken) {
    showMainFlow();
    // Die Absender-Suche stößt applyTargetType() (in setupTargetUi) bereits an.
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

function joinAddresses(arr) {
  return (arr || []).map(formatAddress).join(", ");
}

function getAttachmentNames() {
  var atts = Office.context.mailbox.item.attachments || [];
  return atts.filter(function (a) { return a && !a.isInline; })
             .map(function (a) { return a.name; });
}

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
      applyTargetType();
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
      bodyText = res.value || "";
      el.textContent = bodyText.trim() ? bodyText : "(leerer Text)";
    } else {
      el.textContent = "Text konnte nicht gelesen werden: " +
        (res.error && res.error.message ? res.error.message : "unbekannter Fehler");
    }
  });
}

function autoSearchSender() {
  if (targetType !== "contact") { return; }
  var item = Office.context.mailbox.item;
  var fromEmail = item.from && item.from.emailAddress;
  if (fromEmail) {
    document.getElementById("target-search").value = fromEmail;
    doTargetSearch(fromEmail);
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

  emlSupported = Office.context.requirements.isSetSupported("Mailbox", "1.14");
  if (!emlSupported) {
    emlRadio.disabled = true;
    emlLabel.classList.add("is-disabled");
    emlLabel.title = "Dieser Outlook-Client ist zu alt für den .eml-Export (benötigt Mailbox 1.14).";
    emlLabel.appendChild(document.createTextNode("  (in diesem Client nicht verfügbar)"));
  }

  function refresh() {
    scopeBox.style.display = textRadio.checked ? "block" : "none";
    updateSummary();
    updateSendButton();
  }

  [textRadio, emlRadio, scopeLast, scopeAll].forEach(function (r) {
    if (r) { r.addEventListener("change", refresh); }
  });

  document.getElementById("btn-send").addEventListener("click", onSend);

  document.getElementById("btn-cancel").addEventListener("click", function () {
    textRadio.checked = true;
    scopeLast.checked = true;
    clearSendResult();
    refresh();
  });

  refresh();
}

function updateSummary() {
  var textMode = document.getElementById("mode-text").checked;
  var cfg = TARGET_TYPES[targetType];
  var parts = ["Aktion: " + (textMode ? "Text in Chatter" : "E-Mail (.eml) anhängen")];
  if (textMode) {
    parts.push(document.getElementById("scope-last").checked ? "Nur letzte Nachricht" : "Ganzer Verlauf");
  }
  if (selectedTarget) {
    var label = cfg.label + ": " + selectedTarget.name;
    if (targetType === "task" && selectedProject) {
      label = "Projekt: " + selectedProject.name + " · " + label;
    }
    parts.push(label);
  } else {
    parts.push("Kein Ziel gewählt (" + cfg.label + ")");
  }
  setText("selection-summary", parts.join(" · "));
}

function updateSendButton() {
  var textMode = document.getElementById("mode-text").checked;
  var emlMode = document.getElementById("mode-eml").checked;
  var ok = !!selectedTarget && (textMode || (emlMode && emlSupported));
  document.getElementById("btn-send").disabled = !ok;
}

/* ---------- Zielsuche (Relay -> Odoo) ---------- */

function setupTargetUi() {
  var typeSelect = document.getElementById("target-type");
  var targetInput = document.getElementById("target-search");
  var projectInput = document.getElementById("project-search");

  typeSelect.addEventListener("change", function () {
    targetType = typeSelect.value;
    resetProject();
    resetTarget();
    applyTargetType();
  });

  targetInput.addEventListener("input", function () {
    clearTimeout(searchTimer);
    var q = targetInput.value.trim();
    var min = TARGET_TYPES[targetType].minQuery;
    if (q.length < min) {
      document.getElementById("target-results").innerHTML = "";
      setTargetStatus("");
      return;
    }
    setTargetStatus("Suche…");
    searchTimer = setTimeout(function () { doTargetSearch(q); }, 350);
  });

  projectInput.addEventListener("input", function () {
    clearTimeout(projectSearchTimer);
    var q = projectInput.value.trim();
    setProjectStatus("Suche…");
    projectSearchTimer = setTimeout(function () { doProjectSearch(q); }, 350);
  });

  document.getElementById("sel-change").addEventListener("click", function () {
    resetTarget();
    clearSendResult();
    updateSummary();
    updateSendButton();
    if (TARGET_TYPES[targetType].minQuery === 0 && (targetType !== "task" || selectedProject)) {
      doTargetSearch(document.getElementById("target-search").value.trim());
    }
  });

  document.getElementById("project-sel-change").addEventListener("click", function () {
    resetProject();
    resetTarget();
    clearSendResult();
    updateSummary();
    updateSendButton();
    doProjectSearch(document.getElementById("project-search").value.trim());
  });

  applyTargetType();
}

/* Oberfläche an den gewählten Zieltyp anpassen */
function applyTargetType() {
  var cfg = TARGET_TYPES[targetType];
  var isTask = targetType === "task";

  document.getElementById("project-step").style.display = isTask ? "block" : "none";
  setText("target-label", cfg.label);
  document.getElementById("target-search").placeholder = cfg.placeholder;

  // Bei "Aufgabe" erst nach Projektwahl den Aufgaben-Bereich zeigen
  document.getElementById("target-area").style.display = (isTask && !selectedProject) ? "none" : "block";

  clearSendResult();
  updateSummary();
  updateSendButton();

  if (targetType === "contact") {
    autoSearchSender();
  } else if (targetType === "todo") {
    doTargetSearch("");            // ToDos sofort auflisten
  } else if (isTask) {
    doProjectSearch("");           // Projekte sofort auflisten
  }
}

function resetTarget() {
  selectedTarget = null;
  document.getElementById("target-selected").style.display = "none";
  document.getElementById("target-picker").style.display = "block";
  document.getElementById("target-search").value = "";
  renderList("target-results", "target-status", [], selectTarget);
  setTargetStatus("");
}

function resetProject() {
  selectedProject = null;
  document.getElementById("project-selected").style.display = "none";
  document.getElementById("project-picker").style.display = "block";
  document.getElementById("project-search").value = "";
  renderList("project-results", "project-status", [], selectProject);
  setProjectStatus("");
  if (targetType === "task") {
    document.getElementById("target-area").style.display = "none";
  }
}

function setTargetStatus(msg) { setText("target-status", msg); }
function setProjectStatus(msg) { setText("project-status", msg); }

function searchRelay(payload, onDone, onStatus) {
  if (!clientToken) {
    onStatus("Bitte zuerst den Zugriffs-Token in den Einstellungen eingeben.");
    showSettings();
    return;
  }
  fetch(RELAY_BASE_URL + "/targets/search", {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Client-Token": clientToken },
    body: JSON.stringify(payload)
  }).then(function (res) {
    if (res.status === 401) { throw new Error("401"); }
    if (!res.ok) {
      return res.text().then(function (t) { throw new Error("Relay " + res.status + ": " + t.slice(0, 200)); });
    }
    return res.json();
  }).then(function (data) {
    onDone((data && data.results) || []);
  }).catch(function (err) {
    onDone(null);
    if (err.message === "401") {
      onStatus("Zugriffs-Token ungültig – bitte in den Einstellungen prüfen.");
      showSettings();
    } else {
      onStatus("Fehler bei der Suche: " + err.message);
    }
  });
}

function doTargetSearch(query) {
  var payload = { type: targetType, query: query || "" };
  if (targetType === "task") {
    if (!selectedProject) { return; }
    payload.project_id = selectedProject.id;
  }
  searchRelay(payload, function (results) {
    if (results) { renderList("target-results", "target-status", results, selectTarget); }
  }, setTargetStatus);
}

function doProjectSearch(query) {
  searchRelay({ type: "project", query: query || "" }, function (results) {
    if (results) { renderList("project-results", "project-status", results, selectProject); }
  }, setProjectStatus);
}

function renderList(ulId, statusId, list, onPick) {
  var ul = document.getElementById(ulId);
  ul.innerHTML = "";
  if (!list.length) {
    if (!document.getElementById(statusId).textContent) {
      setText(statusId, "Keine Treffer");
    }
    return;
  }
  setText(statusId, list.length + " Treffer");
  list.forEach(function (r) {
    var li = document.createElement("li");
    li.className = "results__item";

    var nameDiv = document.createElement("div");
    nameDiv.className = "results__name";
    nameDiv.textContent = r.name;

    var metaDiv = document.createElement("div");
    metaDiv.className = "results__meta";
    metaDiv.textContent = r.meta || "—";

    li.appendChild(nameDiv);
    li.appendChild(metaDiv);
    li.addEventListener("click", function () { onPick(r); });
    ul.appendChild(li);
  });
}

function selectProject(p) {
  selectedProject = p;
  document.getElementById("project-picker").style.display = "none";
  document.getElementById("project-selected").style.display = "flex";
  setText("project-sel-name", p.name);
  setText("project-sel-meta", p.meta || "—");
  document.getElementById("target-area").style.display = "block";
  resetTarget();
  clearSendResult();
  updateSummary();
  updateSendButton();
  doTargetSearch("");              // Aufgaben des Projekts sofort auflisten
}

function selectTarget(r) {
  selectedTarget = r;
  document.getElementById("target-picker").style.display = "none";
  document.getElementById("target-selected").style.display = "flex";
  setText("sel-name", r.name);
  setText("sel-meta", r.meta || "—");
  clearSendResult();
  updateSummary();
  updateSendButton();
}

/* ---------- Senden ---------- */

function onSend() {
  if (document.getElementById("mode-eml").checked) {
    sendEml();
  } else {
    sendText();
  }
}

/* Zielangabe für das Relay (res_model + res_id; bei Kontakten zusätzlich
 * partner_id, damit ein noch nicht aktualisiertes Relay weiter funktioniert). */
function targetFields() {
  var fields = {
    res_model: TARGET_TYPES[targetType].model,
    res_id: selectedTarget.id
  };
  if (targetType === "contact") { fields.partner_id = selectedTarget.id; }
  return fields;
}

function sendText() {
  if (!selectedTarget) { return; }
  var item = Office.context.mailbox.item;

  document.getElementById("btn-send").disabled = true;
  setSendResult("sending", "Wird an Odoo gesendet…");

  var scope = document.getElementById("scope-last").checked ? "last" : "all";
  var fields = targetFields();
  var payload = {
    res_model: fields.res_model,
    res_id: fields.res_id,
    partner_id: fields.partner_id,
    scope: scope,
    body_text: bodyText || "",
    meta: {
      subject: item.subject || "",
      sender: formatAddress(item.from),
      to: joinAddresses(item.to),
      cc: joinAddresses(item.cc),
      date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("de-DE") : ""
    },
    attachments: getAttachmentNames()
  };

  postToRelay("/chatter/note", payload);
}

function buildEmlFilename(item) {
  var d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
  var datePart = d.toISOString().slice(0, 10); // YYYY-MM-DD
  var subj = (item.subject || "E-Mail")
    .replace(/[\/\\:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  if (!subj) { subj = "E-Mail"; }
  return datePart + "_" + subj + ".eml";
}

function sendEml() {
  if (!selectedTarget) { return; }
  var item = Office.context.mailbox.item;

  document.getElementById("btn-send").disabled = true;
  setSendResult("sending", "E-Mail wird abgerufen…");

  item.getAsFileAsync(function (fileRes) {
    if (fileRes.status !== Office.AsyncResultStatus.Succeeded) {
      setSendResult("error", "Konnte .eml nicht lesen: " +
        (fileRes.error && fileRes.error.message ? fileRes.error.message : "unbekannter Fehler"));
      updateSendButton();
      return;
    }
    setSendResult("sending", "Wird an Odoo gesendet…");
    var fields = targetFields();
    postToRelay("/chatter/eml", {
      res_model: fields.res_model,
      res_id: fields.res_id,
      partner_id: fields.partner_id,
      filename: buildEmlFilename(item),
      eml_base64: fileRes.value,
      subject: item.subject || ""
    });
  });
}

function postToRelay(path, payload) {
  fetch(RELAY_BASE_URL + path, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Client-Token": clientToken },
    body: JSON.stringify(payload)
  }).then(function (r) {
    if (r.status === 401) { throw new Error("401"); }
    if (!r.ok) {
      return r.text().then(function (t) { throw new Error("Relay " + r.status + ": " + t.slice(0, 300)); });
    }
    return r.json();
  }).then(function (data) {
    setSendResult("ok", "Erfolgreich gesendet ✓", data && data.partner_url);
    updateSendButton();
  }).catch(function (err) {
    if (err.message === "401") {
      setSendResult("error", "Zugriffs-Token ungültig – bitte in den Einstellungen prüfen.");
      showSettings();
    } else {
      setSendResult("error", "Fehler beim Senden: " + err.message);
    }
    updateSendButton();
  });
}

/* ---------- Sende-Ergebnis ---------- */

function clearSendResult() {
  var el = document.getElementById("send-result");
  el.style.display = "none";
  el.textContent = "";
}

function setSendResult(state, message, url) {
  var el = document.getElementById("send-result");
  el.style.display = "block";
  el.className = "send-result send-result--" + state;
  el.textContent = message;
  if (url) {
    el.appendChild(document.createTextNode("  "));
    var a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener";
    a.className = "send-result__link";
    a.textContent = "In Odoo öffnen";
    el.appendChild(a);
  }
}
