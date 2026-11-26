/** stepAheadRobotics - Student Enquiry collector
 * Create/open your Google Sheet, then Extensions > Apps Script and paste this code.
 */
function doPost(e) {
  const sheetName = 'Student Enquiries';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const headers = [
    'Enquiry Date', 'Student Name', 'Student Age', 'Parent Name',
    'Parent WhatsApp', 'Parent Email', 'Interested Course', 'Preferred Level',
    'Country / Time Zone', 'Message'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  const p = (e && e.parameter) ? e.parameter : {};
  sheet.appendRow([
    new Date(), p.studentName || '', p.studentAge || '', p.parentName || '',
    p.parentWhatsApp || '', p.parentEmail || '', p.course || '', p.level || '',
    p.countryTimezone || '', p.message || ''
  ]);

  return ContentService.createTextOutput('OK');
}
