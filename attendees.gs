

// Configuration Constants
const CONFIG = {
  attendeesColumnIndex: 5, // Attendees
  startDateColumnIndex: 6, // Start Date and Time
  endDateColumnIndex: 7, // End Date and Time
  unitsColumn: 'H',
  minutesPerUnitColumn: 'I',
  totalTidColumn: 'J',
  startRow: 2,
  endRow: 100,
  conflictColor: "#800080", // Purple
  dataCheckColumns: [8, 9], // Columns that trigger calculateDuration
};

function checkAttendeeConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const teachersColumn = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  const teachers = teachersColumn.map(row => row[0].trim());

  range.setBackground(null);
  const attendeesToCheck = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const startDate = new Date(row[CONFIG.startDateColumnIndex - 1]);
    const endDate = new Date(row[CONFIG.endDateColumnIndex - 1]);
    const attendees = row[CONFIG.attendeesColumnIndex - 1].split(',');

    const allAttendeesValid = attendees.every(attendee => teachers.includes(attendee.trim()));
    if (!allAttendeesValid) {
      sheet.getRange(i + 1, CONFIG.attendeesColumnIndex).setBackground('red');
    }

    attendees.forEach(attendee => {
      attendee = attendee.trim();
      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + 1 }];
      } else {
        const conflicts = attendeesToCheck[attendee].some(event => {
          const hasConflict = (startDate <= event.end && endDate >= event.start);
          if (hasConflict) {
            sheet.getRange(event.row, CONFIG.attendeesColumnIndex).setBackground('orange');
          }
          return hasConflict;
        });

        if (conflicts) {
          sheet.getRange(i + 1, CONFIG.attendeesColumnIndex).setBackground('orange');
        }
        attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + 1 });
      }
    });
  }
}

function onEdit(e) {


  if ([CONFIG.attendeesColumnIndex, CONFIG.startDateColumnIndex, CONFIG.endDateColumnIndex].includes(e.range.columnStart)) {
    checkAttendeeConflicts();
  }

  if (CONFIG.dataCheckColumns.includes(e.range.columnStart)) {
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