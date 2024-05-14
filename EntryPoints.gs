function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🪄EASV Koordinator Powertools🪄')
    .addItem('Udfyld resten af datoerne for mig⚡', 'generateDates')
    .addItem('Lav skemaer ud fra Moodle grupper (CSV i kolonne M)✨', 'createStudentSchedule')
    .addToUi();
}



function onEdit(e) {

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();


validateEmptyCells(sheet);
validateDateOrder(sheet);
validateLastDate(sheet);
validateNotAllowedDates(sheet);

validateAttendeeTypos(sheet);
validateHoldTypo(sheet);

validateHoldConflicts(sheet);
validateHoldConflicts(sheet);

calculateDuration(sheet);
calculateTotalWeight(sheet);


}