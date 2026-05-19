// =====================================================
// 3DSource Org Chart - Google Apps Script API
// =====================================================
// Deploy this as a Web App from your Google Sheet.
//
// SETUP:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click Deploy > New deployment
// 5. Type: Web app
// 6. Execute as: Me
// 7. Who has access: Anyone (or Anyone within your org)
// 8. Click Deploy and copy the Web App URL
// 9. Paste that URL into index.html as APPS_SCRIPT_URL
// =====================================================

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();

    // First row is headers
    var headers = data[0];
    var employees = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      employees.push({
        id: row[0],
        name: row[1],
        role: row[2],
        department: row[3],
        managerId: row[4] || '',
        location: row[5] || '',
        start: row[6] || '',
        years: row[7] || ''
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, employees: employees }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

