// =============================================
//   Google Apps Script - Backend לאישורי הגעה
// =============================================
//
//   הוראות התקנה:
//
//   1. היכנסו ל-Google Sheets וצרו גיליון חדש
//   2. בשורה הראשונה כתבו את הכותרות:
//      A1: צד  |  B1: שם  |  C1: טלפון  |  D1: מנות  |  E1: סטטוס  |  F1: תאריך
//   3. לחצו על Extensions > Apps Script
//   4. מחקו את הקוד שמופיע והדביקו את כל הקוד הזה
//   5. לחצו Save (Ctrl+S)
//   6. לחצו Deploy > New deployment
//   7. בחרו Type: Web app
//   8. הגדירות:
//      - Execute as: Me
//      - Who has access: Anyone
//   9. לחצו Deploy
//   10. העתיקו את ה-URL שמופיע
//   11. הדביקו את ה-URL בקובץ config.js בשדה SCRIPT_URL
//
//   זהו! המערכת מחוברת.
//
// =============================================

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var action = (e.parameter && e.parameter.action) || 'list';

  if (action === 'add') {
    return addRSVP(sheet, e.parameter);
  } else if (action === 'delete') {
    return deleteRSVP(sheet, e.parameter);
  } else {
    return listRSVP(sheet);
  }
}

function addRSVP(sheet, params) {
  var side = params.side === 'groom' ? 'חתן' : 'כלה';
  var name = params.name || '';
  var phone = params.phone || '';
  var plates = parseInt(params.plates) || 0;
  var status;
  if (params.status === 'coming') {
    status = 'מגיע';
  } else if (params.status === 'not-coming') {
    status = 'לא מגיע';
  } else {
    status = 'עדיין בודק';
  }
  var now = new Date();
  var date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  sheet.appendRow([side, name, phone, plates, status, date]);

  return jsonResponse({ success: true });
}

function deleteRSVP(sheet, params) {
  var row = parseInt(params.row);
  if (row > 1 && row <= sheet.getLastRow()) {
    sheet.deleteRow(row);
    return jsonResponse({ success: true });
  }
  return jsonResponse({ success: false, error: 'Invalid row' });
}

function listRSVP(sheet) {
  var lastRow = sheet.getLastRow();
  var data = [];

  if (lastRow > 1) {
    var range = sheet.getRange(2, 1, lastRow - 1, 6);
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
      data.push({
        side: values[i][0],
        name: values[i][1],
        phone: values[i][2],
        plates: values[i][3],
        status: values[i][4],
        date: values[i][5]
      });
    }
  }

  return jsonResponse({ success: true, data: data });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
