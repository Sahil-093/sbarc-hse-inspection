/**
 * EQUIPMENT INSPECTION
 * Writes to the separate equipment register spreadsheet.
 * Failed checks also open an item on the main Action_Tracker.
 */
const EQUIPMENT_SPREADSHEET_ID = "12M-quo3ISq6KNg6c2jjtFW30EjuwofUJ-0jkQqtcLPk";
const EQUIPMENT_SHEET_GID = 1198663290;

function recordEquipmentInspection(formData, user) {
  try {
    const equipSs = SpreadsheetApp.openById(EQUIPMENT_SPREADSHEET_ID);
    const equipSheet = equipSs.getSheets().find(function (sheet) {
      return sheet.getSheetId() === EQUIPMENT_SHEET_GID;
    });
    if (!equipSheet) {
      return { success: false, message: "Equipment sheet tab was not found." };
    }

    const now = new Date();
    const timeFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "M/d/yyyy H:mm:ss");
    const dateFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

    equipSheet.appendRow([
      timeFormatted,
      formData.contractor,
      formData.equipmentType,
      formData.brand,
      formData.plate,
      formData.equipThirdParty,
      formData.stickerNo,
      formatUsDate_(formData.equipIssue),
      formatUsDate_(formData.equipExpiry),
      formData.operatorName,
      formData.operatorId,
      formData.licence,
      formatUsDate_(formData.licenceExpiry),
      "",
      formData.operatorThirdParty,
      formatUsDate_(formData.operatorIssue),
      formatUsDate_(formData.operatorExpiry),
      formData.herc,
      formData.insurance,
      formData.contact,
      formData.inspectedBy || (user && user.name) || "",
      formData.remarks
    ]);

    const trackerSheet = SPREADSHEET.getSheetByName("Action_Tracker");
    let lastRow = trackerSheet.getLastRow();
    const place = [formData.contractor, formData.equipmentType, formData.plate].filter(Boolean).join(" / ");
    const checks = [
      ["3rd Party", formData.equipThirdParty],
      ["Licence", formData.licence],
      ["Operator 3rd Party", formData.operatorThirdParty],
      ["HERC", formData.herc],
      ["Insurance", formData.insurance]
    ];

    checks.forEach(function (check) {
      if (!isEquipmentFail_(check[1])) return;
      lastRow++;
      trackerSheet.appendRow([
        "ACT-" + (1000 + lastRow),
        dateFormatted,
        formData.zone,
        place,
        "Equipment " + check[0],
        String(check[1] || "") + (formData.remarks ? ". " + formData.remarks : ""),
        formData.targetDate || "",
        "Open",
        "",
        ""
      ]);
    });

    return { success: true, message: "Equipment inspection recorded successfully!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function formatUsDate_(iso) {
  if (!iso) return "";
  const parts = String(iso).split("-");
  if (parts.length !== 3) return String(iso);
  return Number(parts[1]) + "/" + Number(parts[2]) + "/" + parts[0];
}

function isEquipmentFail_(status) {
  const value = String(status || "").trim().toLowerCase();
  return value === "expired" || value === "not available" || value === "no";
}
