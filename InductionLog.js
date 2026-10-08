/**
 * INDUCTION LOG
 * Search by employee name or ID, and append new inductions.
 */
function getInductionSheet_() {
  return SPREADSHEET.getSheetByName("Induction_Log");
}

function ensureInductionSheet_() {
  let sheet = getInductionSheet_();
  if (sheet) return sheet;
  sheet = SPREADSHEET.insertSheet("Induction_Log");
  const headers = ["Timestamp", "Employee Name", "ID Number", "Company", "Trade", "Induction Date", "Expiry Date", "Inducted By", "Zone", "Remarks"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat("@");
  return sheet;
}

function inductionIdText_(value) {
  if (typeof value === "number" && isFinite(value)) return String(Math.round(value));
  return String(value || "").trim();
}

function inductionStatus_(expiry) {
  const iso = cellIso_(expiry);
  if (!iso) return "";
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return iso < today ? "Expired" : "Valid";
}

function inductionRecord_(row, sheetRow) {
  return {
    sheetRow: sheetRow,
    when: formatCell_(row[0]),
    name: String(row[1] || "").trim(),
    idNumber: inductionIdText_(row[2]),
    company: String(row[3] || "").trim(),
    trade: String(row[4] || "").trim(),
    inductionDate: formatCell_(row[5]),
    expiryDate: formatCell_(row[6]),
    inductedBy: String(row[7] || "").trim(),
    zone: String(row[8] || "").trim(),
    remarks: String(row[9] || "").trim(),
    status: inductionStatus_(row[6])
  };
}

function searchInductionLog(query) {
  const q = String(query || "").trim().toLowerCase();
  if (q.length < 2) return [];
  const sheet = getInductionSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getDataRange().getValues();
  const qDigits = q.replace(/\D/g, "");
  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim().toLowerCase();
    const id = inductionIdText_(rows[i][2]).toLowerCase();
    const nameHit = name.indexOf(q) !== -1;
    const idHit = id.indexOf(q) !== -1 || (qDigits.length >= 2 && id.replace(/\D/g, "").indexOf(qDigits) !== -1);
    if (!nameHit && !idHit) continue;
    records.push(inductionRecord_(rows[i], i + 1));
  }
  records.reverse();
  return records.slice(0, 30);
}

function recordInduction(formData, user) {
  try {
    const name = String(formData.name || "").trim();
    const idNumber = String(formData.idNumber || "").trim();
    if (!name || !idNumber) {
      return { success: false, message: "Employee name and ID number are required." };
    }
    const sheet = ensureInductionSheet_();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    sheet.appendRow([
      now,
      name,
      "'" + idNumber,
      formData.company || "",
      formData.trade || "",
      formatUsDate_(formData.inductionDate),
      formatUsDate_(formData.expiryDate),
      formData.inductedBy || (user && user.name) || "",
      formData.zone || "",
      formData.remarks || ""
    ]);
    const last = sheet.getLastRow();
    sheet.getRange(last, 3).setNumberFormat("@").setValue(idNumber);
    return { success: true, message: "Induction saved." };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}
