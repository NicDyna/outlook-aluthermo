/* "An Odoo senden" – Taskpane
 *
 * M1: bestätigt nur, dass Office.js im Client korrekt geladen wurde
 * und das Add-in in Outlook läuft. Noch keine Odoo-Logik.
 */
Office.onReady(function (info) {
  var statusEl = document.getElementById("status");
  var textEl = document.getElementById("status-text");

  function setState(state, message) {
    statusEl.classList.remove("status--loading", "status--ok", "status--error");
    statusEl.classList.add("status--" + state);
    textEl.textContent = message;
  }

  if (info && info.host === Office.HostType.Outlook) {
    // Läuft in Outlook (Web, Neu oder Klassisch)
    setState("ok", "Add-in bereit ✓ (Outlook erkannt)");
  } else if (!info || !info.host) {
    // Direkt im Browser geöffnet – nur zum Prüfen, dass die Seite lädt
    setState("ok", "Vorschau im Browser ✓ – Code geladen (außerhalb von Outlook)");
  } else {
    setState("error", "Unerwartete Umgebung: " + info.host);
  }
});
