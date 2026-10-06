/**
 * GENERATOR INSPECTION MODULE
 * Handles generator data collection and logs automatic action items
 */
function recordGeneratorInspection(formData, user) {
  try {
    const respSheet = SPREADSHEET.getSheetByName("Form Responses 1");
    const trackerSheet = SPREADSHEET.getSheetByName("Action_Tracker");
    const now = new Date();
    const timeFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    const dateFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 1. Raw log save karna
    respSheet.appendRow([
      timeFormatted,
      formData.zone,
      formData.cluster,
      formData.count,
      formData.voltage,
      formData.dripTray,
      formData.dripRemarks,
      formData.fe,
      formData.feRemarks,
      formData.db,
      formData.dbRemarks,
      user.name
    ]);

    // 2. Action items check karna aur Action_Tracker me push karna
    let lastRow = trackerSheet.getLastRow();

    if (formData.dripTray === "No") {
      lastRow++;
      trackerSheet.appendRow([
        "ACT-" + (1000 + lastRow), dateFormatted, formData.zone, formData.cluster, 
        "Drip Tray", formData.dripRemarks || "Drip tray required", 
        formData.targetDate, "Open", "", ""
      ]);
    }
    if (formData.fe === "No") {
      lastRow++;
      trackerSheet.appendRow([
        "ACT-" + (1000 + lastRow), dateFormatted, formData.zone, formData.cluster, 
        "Fire Extinguisher", formData.feRemarks || "Fire Extinguisher missing/discharged", 
        formData.targetDate, "Open", "", ""
      ]);
    }
    if (formData.db === "No") {
      lastRow++;
      trackerSheet.appendRow([
        "ACT-" + (1000 + lastRow), dateFormatted, formData.zone, formData.cluster, 
        "DB Panel", formData.dbRemarks || "DB panel required", 
        formData.targetDate, "Open", "", ""
      ]);
    }

    return { success: true, message: "Generator inspection recorded successfully!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function sheetStatus_(value) {
  const text = String(value || "").trim().toLowerCase();
  if (text === "yes" || text === "valid") return "Valid";
  if (text === "no" || text === "not available") return "Not Available";
  if (text === "expired") return "Expired";
  return String(value || "").trim();
}

function getGeneratorRegister() {
  try {
    const sheet = SPREADSHEET.getSheetByName("Form Responses 1");
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const zone = String(rows[i][1] || "").trim();
      const cluster = String(rows[i][2] || "").trim();
      if (!zone && !cluster) continue;
      records.push({
        category: "Generator",
        sheetRow: i + 1,
        title: cluster || "Generator",
        subtitle: zone,
        when: formatCell_(rows[i][0]),
        inspector: String(rows[i][11] || "").trim(),
        notes: [rows[i][6], rows[i][8], rows[i][10]].map(function (item) {
          return String(item || "").trim();
        }).filter(Boolean).join(" · "),
        statuses: [
          { label: "Drip Tray", value: sheetStatus_(rows[i][5]) },
          { label: "Fire Extinguisher", value: sheetStatus_(rows[i][7]) },
          { label: "DB Panel", value: sheetStatus_(rows[i][9]) }
        ]
      });
    }
    records.reverse();
    return records;
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getScaffoldRegister() {
  try {
    const names = ["Scaffolding", "Scaffold", "Scaffold Inspections", "Scaffolding Inspections"];
    let sheet = null;
    for (let i = 0; i < names.length; i++) {
      sheet = SPREADSHEET.getSheetByName(names[i]);
      if (sheet) break;
    }
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return [];
    const headers = rows[0].map(function (header) { return String(header || "").trim(); });
    const records = [];
    for (let r = 1; r < rows.length; r++) {
      const statuses = [];
      let title = "";
      let subtitle = "";
      headers.forEach(function (header, c) {
        if (!header) return;
        const raw = formatCell_(rows[r][c]);
        if (!title && /scaffold|type|item|location/i.test(header) && raw) title = raw;
        if (!subtitle && /zone|area/i.test(header) && raw) subtitle = raw;
        if (/status|available|expir|valid/i.test(header)) {
          statuses.push({ label: header, value: sheetStatus_(raw) });
        }
      });
      if (!title && !subtitle && !statuses.length) continue;
      records.push({
        category: "Scaffolding",
        sheetRow: r + 1,
        title: title || "Scaffold",
        subtitle: subtitle,
        when: formatCell_(rows[r][0]),
        statuses: statuses
      });
    }
    records.reverse();
    return records;
  } catch (err) {
    return { success: false, message: String(err) };
  }
}