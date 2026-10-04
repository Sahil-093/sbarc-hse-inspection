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