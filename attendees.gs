// Configuration Constants
const CONFIG = {
  attendeesColumn: 'E',
  startDateColumn: 'F',
  endDateColumn: 'G',
  unitsColumn: 'H',
  minutesPerUnitColumn: 'I',
  totalTidColumn: 'J',
  startRow: 2,
  endRow: 100,
  conflictColor: "#800080", // Purple
  dataCheckColumns: ['H', 'I'], // Columns that trigger calculateDuration
};

function checkAttendeeTypos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const teachersColumn = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  const teachers = teachersColumn.map(row => row[0].trim());
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  const attendeesValues = attendeesRange.getValues();

  attendeesRange.setBackground(null);

  attendeesValues.forEach((row, i) => {
    const attendees = row[0].trim().split(/,\s*/);
    const allAttendeesValid = attendees.every(attendee => teachers.includes(attendee.trim()));
    if (!allAttendeesValid) {
      sheet.getRange(CONFIG.attendeesColumn + (i + CONFIG.startRow)).setBackground('red');
    }
  });
}

function checkAttendeeConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const attendeesToCheck = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const startDate = new Date(row[sheet.getRange(CONFIG.startDateColumn + '1').getColumn() - 1]);
    const endDate = new Date(row[sheet.getRange(CONFIG.endDateColumn + '1').getColumn() - 1]);
    const attendees = row[sheet.getRange(CONFIG.attendeesColumn + '1').getColumn() - 1].trim().split(/,\s*/);

    attendees.forEach(attendee => {
      attendee = attendee.trim();
      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + CONFIG.startRow }];
      } else {
        const conflicts = attendeesToCheck[attendee].some(event => {
          const hasConflict = (startDate <= event.end && endDate >= event.start);
          if (hasConflict) {
            sheet.getRange(event.row, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground('orange');
          }
          return hasConflict;
        });

        if (conflicts) {
          sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground('orange');
        }
        attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + CONFIG.startRow });
      }
    });
  }
}

function onEdit(e) {
  if ([CONFIG.attendeesColumn, CONFIG.startDateColumn, CONFIG.endDateColumn].includes(e.range.getA1Notation().charAt(0))) {
    checkAttendeeTypos();
    checkAttendeeConflicts();
  }

  if (CONFIG.dataCheckColumns.includes(e.range.getA1Notation().charAt(0))) {
    calculateDuration();
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