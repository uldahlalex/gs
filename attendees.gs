
const CONFIG = {
  possibleAttendeesColumn: 'A',
  attendeesColumn: 'E', // Attendees, originally column index 5
  startDateColumn: 'F', // Start Date and Time, originally column index 6
  endDateColumn: 'G', // End Date and Time, originally column index 7
  unitsColumn: 'H',
  minutesPerUnitColumn: 'I',
  totalTidColumn: 'J',
  startRow: 2,
  endRow: 100,
  conflictColor: "#800080", // Purple
  dataCheckColumns: [CONFIG.unitsColumn, CONFIG.minutesPerUnitColumn], // Columns that trigger calculateDuration
};

function columnToIndex(columnLetter) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnLetter + '1').getColumn();
}

function onEdit(e) {
  const editedColumn = e.range.getColumn();
  if ([columnToIndex(CONFIG.attendeesColumn), columnToIndex(CONFIG.startDateColumn), columnToIndex(CONFIG.endDateColumn)].includes(editedColumn)) {
    checkAttendeeConflicts();
  }

  if (CONFIG.dataCheckColumns.map(columnToIndex).includes(editedColumn)) {
    calculateDuration();
  }
}
function checkAttendeeConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  const possibleAttendeesColumn = sheet.getRange(CONFIG.startRow, columnToIndex(CONFIG.possibleAttendeesColumn), sheet.getLastRow() - CONFIG.startRow + 1).getValues();
  const teachers = possibleAttendeesColumn.map(row => row[0].trim());

  // Reset background for attendees column only
  sheet.getRange(CONFIG.startRow, columnToIndex(CONFIG.attendeesColumn), sheet.getLastRow() - CONFIG.startRow + 1, 1).setBackground(null);

  const attendeesToCheck = {};

  // Loop through each data row starting from CONFIG.startRow
  for (let i = CONFIG.startRow - 1; i < values.length; i++) {
    const row = values[i];
    const startDate = new Date(row[columnToIndex(CONFIG.startDateColumn) - 1]);
    const endDate = new Date(row[columnToIndex(CONFIG.endDateColumn) - 1]);
    const attendees = row[columnToIndex(CONFIG.attendeesColumn) - 1].split(',');

    let allAttendeesValid = attendees.every(attendee => teachers.includes(attendee.trim()));

    if (!allAttendeesValid) {
      sheet.getRange(i + 1, columnToIndex(CONFIG.attendeesColumn)).setBackground('red');
    }

    attendees.forEach(attendee => {
      attendee = attendee.trim();
      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [];
      }
      attendeesToCheck[attendee].forEach(event => {
        const hasConflict = (startDate <= event.end && endDate >= event.start) || (startDate >= event.start && endDate <= event.end);
        if (hasConflict) {
          sheet.getRange(i + 1, columnToIndex(CONFIG.attendeesColumn)).setBackground(CONFIG.conflictColor);
          sheet.getRange(event.row, columnToIndex(CONFIG.attendeesColumn)).setBackground(CONFIG.conflictColor);
        }
      });
      attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + 1 });
    });
  }
}
function calculateDuration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const numRows = CONFIG.endRow - CONFIG.startRow + 1;
  const unitsRange = sheet.getRange(CONFIG.unitsColumn + CONFIG.startRow + ':' + CONFIG.unitsColumn + CONFIG.endRow);
  const minutesPerUnitRange = sheet.getRange(CONFIG.minutesPerUnitColumn + CONFIG.startRow + ':' + CONFIG.minutesPerUnitColumn + CONFIG.endRow);
  const unitsValues = unitsRange.getValues();
  const minutesPerUnitValues = minutesPerUnitRange.getValues();

  for (let i = 0; i < numRows; i++) {
    const numUnits = unitsValues[i][0];
    const minutesPerUnit = minutesPerUnitValues[i][0];
    if (numUnits && minutesPerUnit) {
      const totalTimeInDays = (numUnits * minutesPerUnit) / 24 / 60;
      const totalTidCell = sheet.getRange(CONFIG.totalTidColumn + (i + CONFIG.startRow));
      totalTidCell.setValue(totalTimeInDays);
      totalTidCell.setNumberFormat('[h]:mm');
    }
  }
}