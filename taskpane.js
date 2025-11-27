// Wird aufgerufen, wenn das Add-In bereit ist
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      // prüfen, ob das Event unterstützt wird
      if (!sheet.onSelectionChanged || !sheet.onSelectionChanged.add) {
        showError("Dieses Excel unterstützt 'onSelectionChanged' nicht.");
        return;
      }

      // Auswahl-Ereignis auf dem aktiven Blatt registrieren
      sheet.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();

      clearPanel("Zellmarkierung in Spalte P oder rechts wählen.");
    }).catch(errorHandler);
  }
});

// Panel leeren oder Hinweistext setzen
function clearPanel(message) {
  const div = document.getElementById("info");
  if (!div) return;

  div.className = "empty-hint";
  div.textContent = message || "Keine Anzeige für diese Zelle.";
}

// Fehler-Text im Panel anzeigen
function showError(message) {
  const div = document.getElementById("info");
  if (!div) return;

  div.className = "empty-hint";
  div.textContent = "Fehler: " + message;
}

// Haupt-Handler bei Auswahländerung
async function handleSelectionChanged(eventArgs) {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      // aktuelle Auswahl
      const range = sheet.getSelectedRange();
      range.load(["columnIndex", "rowIndex", "format/fill/color"]);
      await context.sync();

      const colIndex = range.columnIndex; // 0-basiert (A=0, B=1, ..., P=15)
      const rowIndex = range.rowIndex;

      // 1) Nur ab Spalte P (Index 15) reagieren
      if (colIndex < 15) {
        clearPanel("Zellmarkierung in Spalte P oder rechts wählen.");
        return;
      }

      // 2) Farbfilter: bestimmte Farben sollen NICHT angezeigt werden
      let rawColor = range.format.fill.color || "";
      let normalized = rawColor.toString().replace("#", "").toUpperCase();

      const blockedColors = [
        "FFFFFF", // weiß
        "404040", // dunkelgrau
        "A6A6A6"  // mittelgrau
      ];

      if (blockedColors.includes(normalized)) {
        clearPanel("Für Zellen mit dieser Farbe wird nichts angezeigt.");
        return;
      }

      // 3) Daten der gleichen Zeile aus D:O holen
      const rowNumber = rowIndex + 1; // A1-Notation ist 1-basiert
      const rowRange = sheet.getRange(`D${rowNumber}:O${rowNumber}`);
      rowRange.load("values");
      await context.sync();

      const v = rowRange.values[0];
      // D..O → Indexe:
      // 0=D, 1=E, 2=F, 3=G, 4=H, 5=I, 6=J, 7=K, 8=L, 9=M, 10=N, 11=O

      const start       = v[0];  // D
      const ende        = v[1];  // E
      const dauer       = v[2];  // F
      const fachbereich = v[3];  // G
      const kursnummer  = v[4];  // H
      const kurs        = v[5];  // I
      const teilnehmer  = v[6];  // J
      const kursleiter  = v[7];  // K
      const freigabe    = v[8];  // L
      const rf          = v[9];  // M
      const fahrzeuge   = v[10]; // N
      const anlagen     = v[11]; // O

      const div = document.getElementById("info");
      if (!div) return;

      const esc = (val) =>
        val === null || val === undefined || val === "" ? "-" : val;

      div.className = "";
      div.innerHTML = `
        <div class="info-row"><span class="label">Fachbereich:</span> ${esc(fachbereich)}</div>
        <div class="info-row"><span class="label">Kursnummer:</span> ${esc(kursnummer)}</div>
        <div class="info-row"><span class="label">Kurs:</span> ${esc(kurs)}</div>
        <div class="info-row"><span class="label">Start:</span> ${esc(start)}</div>
        <div class="info-row"><span class="label">Ende:</span> ${esc(ende)}</div>
        <div class="info-row"><span class="label">Dauer:</span> ${esc(dauer)}</div>
        <div class="info-row"><span class="label">Kursleiter:</span> ${esc(kursleiter)}</div>
        <div class="info-row"><span class="label">Teilnehmer:</span> ${esc(teilnehmer)}</div>
        <div class="info-row"><span class="label">RF:</span> ${esc(rf)}</div>
        <div class="info-row"><span class="label">Fahrzeuge:</span> ${esc(fahrzeuge)}</div>
        <div class="info-row"><span class="label">Anlagen:</span> ${esc(anlagen)}</div>
        <div class="info-row"><span class="label">Freigabe:</span> ${esc(freigabe)}</div>
      `;
    });
  } catch (error) {
    errorHandler(error);
  }
}

// Fehlerhandler: schreibt den Fehlertext ins Panel
function errorHandler(error) {
  console.error(error);
  let msg = "Unbekannter Fehler.";

  if (error && error.message) {
    msg = error.message;
  }
  if (error && error.debugInfo && error.debugInfo.errorLocation) {
    msg += " (Ort: " + error.debugInfo.errorLocation + ")";
  }

  showError(msg);
}
