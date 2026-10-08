/**
 * INDUCTION LOG
 * Reads and appends the SBARC induction register.
 */
var INDUCTION_SPREADSHEET_ID = "1BxOtq6QKLEwsSrhiv2cIRZdtNiS5k2pt2uZaFIep4pY";
var INDUCTION_SHEET_GID = 578338714;

function inductionTab_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === INDUCTION_SHEET_GID) return sheets[i];
  }
  return null;
}

function readInductionRows_() {
  try {
    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) throw new Error("Induction sheet tab was not found.");
    return sheet.getDataRange().getValues();
  } catch (err) {
    const url = "https://docs.google.com/spreadsheets/d/" + INDUCTION_SPREADSHEET_ID +
      "/export?format=csv&gid=" + INDUCTION_SHEET_GID;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) throw err;
    return Utilities.parseCsv(response.getContentText());
  }
}

function inductionIdText_(value) {
  if (typeof value === "number" && isFinite(value)) return String(Math.round(value));
  return String(value || "").trim();
}

function inductionHeaderIndex_(rows) {
  const limit = Math.min(rows.length, 6);
  for (let i = 0; i < limit; i++) {
    if (String(rows[i][1] || "").trim().toLowerCase() === "name") return i;
  }
  return 1;
}

function inductionDateText_(row) {
  const day = inductionIdText_(row[5]);
  const month = String(row[6] || "").trim();
  const year = inductionIdText_(row[7]);
  return [day, month, year].filter(Boolean).join(" ");
}

function inductionRecord_(row, sheetRow) {
  return {
    sheetRow: sheetRow,
    serial: inductionIdText_(row[0]),
    name: String(row[1] || "").trim(),
    sticker: String(row[2] || "").trim(),
    idNumber: inductionIdText_(row[3]),
    position: String(row[4] || "").trim(),
    inductionDate: inductionDateText_(row),
    company: String(row[8] || "").trim(),
    trainer: String(row[9] || "").trim(),
    notes: String(row[10] || "").trim()
  };
}

function searchInductionLog(query) {
  const q = String(query || "").trim().toLowerCase();
  if (q.length < 2) return [];
  const rows = readInductionRows_();
  const start = inductionHeaderIndex_(rows) + 1;
  const qDigits = q.replace(/\D/g, "");
  const records = [];
  for (let i = start; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim();
    const id = inductionIdText_(rows[i][3]);
    if (!name && !id) continue;
    const nameHit = name.toLowerCase().indexOf(q) !== -1;
    const idHit = id && (id.toLowerCase().indexOf(q) !== -1 || (qDigits.length >= 2 && id.replace(/\D/g, "").indexOf(qDigits) !== -1));
    if (!nameHit && !idHit) continue;
    records.push(inductionRecord_(rows[i], i + 1));
  }
  records.reverse();
  return records.slice(0, 40);
}

function recordInduction(formData, user) {
  try {
    const name = String(formData.name || "").trim();
    const idNumber = String(formData.idNumber || "").trim();
    if (!name) return { success: false, message: "Employee name is required." };

    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) return { success: false, message: "Induction sheet tab was not found." };

    const parts = String(formData.inductionDate || "").split("-");
    const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const day = parts.length === 3 ? Number(parts[2]) : "";
    const month = parts.length === 3 ? months[Number(parts[1]) - 1] || "" : "";
    const year = parts.length === 3 ? Number(parts[0]) : "";

    const rows = sheet.getDataRange().getValues();
    let serial = 0;
    for (let i = 0; i < rows.length; i++) {
      const n = parseInt(rows[i][0], 10);
      if (!isNaN(n) && n > serial) serial = n;
    }

    sheet.appendRow([
      serial + 1,
      name,
      formData.sticker || "",
      idNumber,
      formData.position || "",
      day,
      month,
      year,
      formData.company || "",
      formData.trainer || (user && user.name) || "",
      formData.notes || ""
    ]);
    const last = sheet.getLastRow();
    if (idNumber) sheet.getRange(last, 4).setNumberFormat("@").setValue(idNumber);
    return { success: true, message: "Induction saved." };
  } catch (err) {
    const message = String(err);
    if (/permission|access/i.test(message)) {
      return { success: false, message: "Cannot write the induction sheet. Share it with sahilkhattak093@gmail.com as Editor." };
    }
    return { success: false, message: message };
  }
}
