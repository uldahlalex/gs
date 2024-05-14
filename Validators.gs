
function validateHoldConflicts(sheet) {
  const holdInterval = parseInt(sheet.getRange(CONFIG.holdInterval).getValue());
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const holdsSchedule = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const holds = row[sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn() - 1].split(',');
    const startDate = new Date(row[sheet.getRange(CONFIG.startDateColumn + '1').getColumn() - 1]);
    const endDate = new Date(row[sheet.getRange(CONFIG.endDateColumn + '1').getColumn() - 1]);

    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    holds.forEach(hold => {
      hold = hold.trim();
      
      if (hold) {
        if (!holdsSchedule[hold]) {
          holdsSchedule[hold] = [];
        }
        
        const conflicts = holdsSchedule[hold].some(event => {
          const adjustedEventEnd = new Date(event.end);
          adjustedEventEnd.setDate(adjustedEventEnd.getDate() + holdInterval);

          return startDate < adjustedEventEnd && endDate > event.start;
        });

        if (conflicts) {
          sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn()).setBackground(CONFIG.dateValidationErrorColor);
        } else {
          holdsSchedule[hold].push({ start: startDate, end: endDate });
        }
      }
    });
  }
}

function validateAttendeeConflicts(sheet) {
  const attendeeInterval = parseInt(sheet.getRange(CONFIG.attendeeInterval).getValue());
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const attendeesToCheck = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const startDate = new Date(row[sheet.getRange(CONFIG.startDateColumn + '1').getColumn() - 1]);
    const endDate = new Date(row[sheet.getRange(CONFIG.endDateColumn + '1').getColumn() - 1]);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);
    const attendees = row[sheet.getRange(CONFIG.attendeesColumn + '1').getColumn() - 1].trim().split(/,\s*/);

    attendees.forEach(attendee => {
      attendee = attendee.trim();
      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + CONFIG.startRow }];
      } else {
        const conflicts = attendeesToCheck[attendee].some(event => {
          const adjustedEventEnd = new Date(event.end);
          adjustedEventEnd.setDate(adjustedEventEnd.getDate() + attendeeInterval);
          const hasConflict = (startDate <= adjustedEventEnd && endDate >= event.start);

          if (hasConflict) {
            sheet.getRange(event.row, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground(CONFIG.dateValidationErrorColor);
            sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground(CONFIG.dateValidationErrorColor);
          }
          return hasConflict;
        });

        if (!conflicts) {
          attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + CONFIG.startRow });
        }
      }
    });
  }

}
function validateAttendeeTypos(sheet) {
  const teachersRange = sheet.getRange(CONFIG.teachersColumn + '2:' + CONFIG.teachersColumn + sheet.getLastRow());
  const teachersValues = teachersRange.getValues();
  const teachers = teachersValues.map(row => row[0].trim());
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  const attendeesValues = attendeesRange.getValues();


  attendeesValues.forEach((row, i) => {
    const attendeeCellContent = row[0].trim();
    if (attendeeCellContent) { 
      const attendees = attendeeCellContent.split(/,\s*/);
      const allAttendeesValid = attendees.every(attendee => teachers.includes(attendee.trim()));
      if (!allAttendeesValid) {
        sheet.getRange(CONFIG.attendeesColumn + (i + CONFIG.startRow)).setBackground(CONFIG.invalidDataColor);
      }
    }
  });
}

function validateDateOrder(sheet) {
  const startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ":" + CONFIG.startDateColumn + CONFIG.endRow);
  const endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ":" + CONFIG.endDateColumn + CONFIG.endRow);
  const startDateValues = startDateRange.getValues().flat();
  const endDateValues = endDateRange.getValues().flat();

  startDateValues.forEach((startDate, i) => {
    const endDate = endDateValues[i];
    if (new Date(startDate) > new Date(endDate)) {
      startDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateValidationErrorColor);
      endDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateValidationErrorColor);
    }
  });
}

function validateNotAllowedDates(sheet) {
  const notAllowedDatesRange = sheet.getRange(CONFIG.notAllowedDates + "2:" + CONFIG.notAllowedDates + (CONFIG.startRow + 15));
  const notAllowedDates = notAllowedDatesRange.getValues().flat().map(date => new Date(date).setHours(0, 0, 0, 0));

  const startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ":" + CONFIG.startDateColumn + CONFIG.endRow);
  const endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ":" + CONFIG.endDateColumn + CONFIG.endRow);
  const startDateValues = startDateRange.getValues().flat();
  const endDateValues = endDateRange.getValues().flat();


  startDateValues.forEach((startDate, i) => {
    const endDate = endDateValues[i];
    notAllowedDates.forEach((notAllowedDate, j) => {
      if (notAllowedDate >= new Date(startDate) && notAllowedDate <= new Date(endDate)) {
        startDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateValidationErrorColor);
        endDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateValidationErrorColor);
      }
    });
  });
}


function validateHoldTypo(sheet) {
  const allowedHoldsRange = sheet.getRange(CONFIG.tilladteHold + '2:' + CONFIG.tilladteHold + sheet.getLastRow());
  const allowedHoldsValues = allowedHoldsRange.getValues();
  const allowedHolds = allowedHoldsValues.map(row => row[0].trim());
  const holdsRange = sheet.getRange(CONFIG.eksamensHoldColumn + CONFIG.startRow + ':' + CONFIG.eksamensHoldColumn + CONFIG.endRow);
  const holdsValues = holdsRange.getValues();

  holdsValues.forEach((row, i) => {
    const holdCellContent = row[0].trim();
    if (holdCellContent) {
      const holds = holdCellContent.split(/,\s*/);
      const allHoldsValid = holds.every(hold => allowedHolds.includes(hold.trim()));
      if (!allHoldsValid) {
        sheet.getRange(CONFIG.eksamensHoldColumn + (i + CONFIG.startRow)).setBackground(CONFIG.invalidDataColor);
      }
    }
  });
}

function validateLastDate(sheet) {
  
  const startDateColIndex = sheet.getRange(CONFIG.startDateColumn + '1').getColumn();

  const latestDate = new Date(sheet.getRange(CONFIG.latestDateCell).getValue());
  
  const numRows = CONFIG.endRow - CONFIG.startRow + 1;
  const dateRange = sheet.getRange(CONFIG.startRow, startDateColIndex, numRows, 2); 
  const dateValues = dateRange.getValues();

  for (let i = 0; i < numRows; i++) {
    const startDate = new Date(dateValues[i][0]);
    const endDate = new Date(dateValues[i][1]); 

    if (startDate > endDate || endDate > latestDate) {
      dateRange.getCell(i + 1, 1).setBackground(CONFIG.dateValidationErrorColor); 
      dateRange.getCell(i + 1, 2).setBackground(CONFIG.dateValidationErrorColor);
    } 
  }
}


function validateEmptyCells(sheet) {
  const examNameColIndex = sheet.getRange(CONFIG.examNameColumn + '1').getColumn();
  const attendeesColIndex = sheet.getRange(CONFIG.attendeesColumn + '1').getColumn();
  const unitsColIndex = sheet.getRange(CONFIG.unitsColumn + '1').getColumn();
  const minutesPerUnitColIndex = sheet.getRange(CONFIG.minutesPerUnitColumn + '1').getColumn();
  const eksamensHoldIndex = sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn();

  const startRow = CONFIG.startRow;
  const endRow = CONFIG.endRow;
  const numRows = endRow - startRow + 1;
  
  const range = sheet.getRange(startRow, 1, numRows, sheet.getMaxColumns());
  const values = range.getValues();

  for (let i = 0; i < numRows; i++) {
    const row = values[i];
    // Only validate rows with a non-empty exam name
    if (row[examNameColIndex - 1].trim()) {
      checkAndColorIfEmpty(row, examNameColIndex, i, sheet, CONFIG.invalidDataColor);
      checkAndColorIfEmpty(row, attendeesColIndex, i, sheet, CONFIG.invalidDataColor);
      checkAndColorIfEmpty(row, unitsColIndex, i, sheet, CONFIG.invalidDataColor);
      checkAndColorIfEmpty(row, minutesPerUnitColIndex, i, sheet, CONFIG.invalidDataColor);
      checkAndColorIfEmpty(row, eksamensHoldIndex, i, sheet, CONFIG.invalidDataColor);
    }
  }
}

function checkAndColorIfEmpty(row, columnIndex, rowIndex, sheet, color) {
  if (!String(row[columnIndex - 1]).trim()) {
    sheet.getRange(CONFIG.startRow + rowIndex, columnIndex).setBackground(color);
  }
}

