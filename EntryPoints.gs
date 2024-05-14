function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🪄EASV Koordinator Powertools🪄')
    .addItem('Udfyld resten af datoerne for mig⚡', 'generateDates')
    .addItem('Lav skemaer ud fra Moodle grupper (CSV i kolonne M)✨', 'createStudentSchedule')
    .addToUi();
}



function onEdit(e) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var totalRows = sheet.getMaxRows();
  var totalColumns = sheet.getMaxColumns();
  var range = sheet.getRange(1, 1, totalRows, totalColumns);
range.setBackground(null);

validateEmptyCells(sheet);
validateNotAllowedDates(sheet);
validateAttendeeTypos(sheet);
validateAttendeeConflicts(sheet);
validateHoldTypo(sheet);
validateHoldConflicts(sheet);
validateDateOrder(sheet);
validateLastDate(sheet);

calculateDuration(sheet);
calculateTotalWeight(sheet);


}