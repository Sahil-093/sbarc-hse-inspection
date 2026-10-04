/**
 * MAIN CONTROLLER
 * Handles Routing, Core Configuration, and Authentication
 */
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

// Web App entry point
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HSE Inspection & Compliance Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Global user authentication (from HSE_TEAM sheet)
function authenticateUser(username, pin) {
  const sheet = SPREADSHEET.getSheetByName("HSE_TEAM");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const uName = String(data[i][2]).trim().toLowerCase();
    const uPin = String(data[i][3]).trim();
    
    if (uName === username.trim().toLowerCase() && uPin === pin.trim()) {
      const assignedZonesStr = String(data[i][4]);
      return {
        success: true,
        user: {
          id: data[i][0],
          name: data[i][1],
          username: data[i][2],
          assignedZones: assignedZonesStr.split(',').map(s => s.trim()),
          role: data[i][5]
        }
      };
    }
  }
  return { success: false, message: "Invalid username or PIN." };
}