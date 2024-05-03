const CONFIG = {
  teachersColumn: 'A',
  notAllowedDates: 'B',
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
  nonEditableColumns: ['J'],

  earliestDateCell: 'B33',
  latestDateCell: 'B34',
};


function checkAttendeeTypos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const teachersRange = sheet.getRange(CONFIG.teachersColumn + '2:' + CONFIG.teachersColumn + sheet.getLastRow());
  const teachersValues = teachersRange.getValues();
  const teachers = teachersValues.map(row => row[0].trim());
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

function checkDateConflictsAndColorCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var notAllowedDatesRange = sheet.getRange(CONFIG.notAllowedDates + CONFIG.startRow + ":" + CONFIG.notAllowedDates + (CONFIG.startRow + 15));
  var notAllowedDates = notAllowedDatesRange.getValues().flat();
  var startDateRange = sheet.getRange(CONFIG.startDateColumn + (CONFIG.startRow + 1) + ":" + CONFIG.startDateColumn);
  var startDateValues = startDateRange.getValues().flat();
  var endDateRange = sheet.getRange(CONFIG.endDateColumn + (CONFIG.startRow + 1) + ":" + CONFIG.endDateColumn);
  var endDateValues = endDateRange.getValues().flat();

  // Clear any previous formatting
  notAllowedDatesRange.setBackground(null);
  startDateRange.setBackground(null);
  endDateRange.setBackground(null);

  // Loop through all start and end dates
  for (var i = 0; i < startDateValues.length; i++) {
    var startDate = startDateValues[i];
    var endDate = endDateValues[i];
    if (startDate && endDate) { // Check if both start and end dates are present
      startDate = new Date(startDate);
      endDate = new Date(endDate);
      // Check against each not allowed date
      for (var j = 0; j < notAllowedDates.length; j++) {
        var notAllowedDate = new Date(notAllowedDates[j]);
        if (notAllowedDate >= startDate && notAllowedDate <= endDate) {
          // If conflict, color the cells
          notAllowedDatesRange.getCell(j + 1, 1).setBackground(CONFIG.conflictColor);
          startDateRange.getCell(i + 1, 1).setBackground(CONFIG.conflictColor);
          endDateRange.getCell(i + 1, 1).setBackground(CONFIG.conflictColor);
        }
      }
    }
  }
}
function onEdit(e) {
  if ([CONFIG.attendeesColumn, CONFIG.startDateColumn, CONFIG.endDateColumn, CONFIG.notAllowedDates].includes(e.range.getA1Notation().charAt(0))) {
    checkAttendeeTypos();
    checkAttendeeConflicts();
    checkDateConflictsAndColorCells();
  }

  if (CONFIG.dataCheckColumns.includes(e.range.getA1Notation().charAt(0))) {
    calculateDuration(e);
  }

  if (CONFIG.nonEditableColumns.includes(e.range.getA1Notation().charAt(0))) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Bemærk: Det er ikke mening "total tid" manuelt skal sættes, da det automatisk sker når de forrige to kolonners værdier ændres');
    return;
  }


  if (e.range.getA1Notation() === "B30" && e.range.getValue() === "Go") {
    generateDates();
  }
}
function generateDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  //const earliestDate = parseDate("6/7/2024");
  const earliestDate = sheet.getRange("B33").getValue();
  //const latestDate = parseDate("1/7/2024")
  const latestDate = sheet.getRange("B34").getValue();
  const attendeesSchedule = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const attendees = row[sheet.getRange(CONFIG.attendeesColumn + '1').getColumn() - 1].trim().split(/,\s*/);
    const totalTime = row[sheet.getRange(CONFIG.totalTidColumn + '1').getColumn() - 1];
    const startDateCell = sheet.getRange(CONFIG.startDateColumn + (i + CONFIG.startRow));
    const endDateCell = sheet.getRange(CONFIG.endDateColumn + (i + CONFIG.startRow));

    if (!startDateCell.getValue() && !endDateCell.getValue() && totalTime) {

      const totalDays = Math.ceil(totalTime / 6); // Convert hours to days, assuming 6 hours per day
      let startDate = earliestDate;
      let endDate = new Date(startDate);

      Browser.msgBox(totalTime, Browser.Buttons.OK_CANCEL);
      endDate.setDate(endDate.getDate() + totalDays - 1);
      Browser.msgBox(endDate, Browser.Buttons.OK_CANCEL);
      while (endDate > latestDate) {
        startDate.setDate(startDate.getDate() + 1);
        endDate.setDate(startDate.getDate() + totalDays - 1);
      }

      let hasConflict = false;
      attendees.forEach(attendee => {
        if (attendeesSchedule[attendee]) {
          hasConflict = attendeesSchedule[attendee].some(event => {
            return (startDate <= event.end && endDate >= event.start);
          });
        }
      });

      if (!hasConflict) {
        startDateCell.setValue(startDate);
        endDateCell.setValue(endDate);
        attendees.forEach(attendee => {
          if (!attendeesSchedule[attendee]) {
            attendeesSchedule[attendee] = [];
          }
          attendeesSchedule[attendee].push({ start: startDate, end: endDate });
        });
      }
    }
  }
}

function parseDate(str) {
  var parts = str.split("/");
  return new Date(parseInt(parts[2], 10),
      parseInt(parts[1], 10) - 1,
      parseInt(parts[0], 10));
}

function formatDate(date, format) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  return format.replace('dd', day).replace('MM', month).replace('yyyy', year).replace('HH', hours).replace('mm', minutes);
}
function calculateDuration(e) {
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
      const totalTimeInHours = (numUnits * minutesPerUnit) / 60;
      const totalTidCell = sheet.getRange(CONFIG.totalTidColumn + (i + CONFIG.startRow));
      totalTidCell.setValue(totalTimeInHours);
      totalTidCell.setNumberFormat('#,##0.00');

      // Check if the total time exceeds 50 hours
      if (totalTimeInHours > 50) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Advarsel om mulig fejl-indtastning: Beregnet tid er over 50 timer, tjek venligst minutter per eksamen og holdstørrelse");
      }
    }
  }
}