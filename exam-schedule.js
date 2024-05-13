const CONFIG = {
  teachersColumn: 'A',
  totalWeightPerEmployeeColumn: 'B',
  tilladteHold: 'C',
  notAllowedDates: 'D',
  examNameColumn: 'E',
  eksamensHoldColumn: 'F',
  attendeesColumn: 'G',
  unitsColumn: 'H',
  minutesPerUnitColumn: 'I',
  totalTidColumn: 'J',
  startDateColumn: 'K',
  endDateColumn: 'L',
  csvDataColumn: 'M',

  maksTid: 'O2',
  interval: 'O3',
  earliestDateCell: 'O4',
  latestDateCell: 'O8',
  holdInterval: 'O9',
  attendeeInterval: 'O10',
  startRow: 2,
  endRow: 100,
  dateConflictColor: 'orange',
  invalidDataColor: 'red'
};


function clearAttendeeConflictColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  attendeesRange.setBackgrounds(attendeesRange.getBackgrounds().map(row => row.map(cell => cell === CONFIG.dateConflictColor ? null : cell)));
}

function clearAttendeeTypoColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  attendeesRange.setBackgrounds(attendeesRange.getBackgrounds().map(row => row.map(cell => cell === CONFIG.invalidDataColor ? null : cell)));
}

function clearHoldConflictColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const holdsRange = sheet.getRange(CONFIG.eksamensHoldColumn + CONFIG.startRow + ':' + CONFIG.eksamensHoldColumn + CONFIG.endRow);
  holdsRange.setBackgrounds(holdsRange.getBackgrounds().map(row => row.map(cell => cell === CONFIG.dateConflictColor ? null : cell)));
}

function clearHoldTypoColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const holdsRange = sheet.getRange(CONFIG.eksamensHoldColumn + CONFIG.startRow + ':' + CONFIG.eksamensHoldColumn + CONFIG.endRow);
  holdsRange.setBackgrounds(holdsRange.getBackgrounds().map(row => row.map(cell => cell === CONFIG.invalidDataColor ? null : cell)));
}

function clearDateConflictColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ':' + CONFIG.startDateColumn + CONFIG.endRow);
  const endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ':' + CONFIG.endDateColumn + CONFIG.endRow);
  
  // Get the backgrounds for the start and end date ranges.
  const startBackgrounds = startDateRange.getBackgrounds();
  const endBackgrounds = endDateRange.getBackgrounds();
  
  // Clear the colors that match CONFIG.dateConflictColor.
  const clearedStartBackgrounds = startBackgrounds.map(row => row.map(cellColor => cellColor === CONFIG.dateConflictColor ? null : cellColor));
  const clearedEndBackgrounds = endBackgrounds.map(row => row.map(cellColor => cellColor === CONFIG.dateConflictColor ? null : cellColor));

  // Set the cleared backgrounds back to the ranges.
  startDateRange.setBackgrounds(clearedStartBackgrounds);
  endDateRange.setBackgrounds(clearedEndBackgrounds);
}

function clearInvalidInputColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const columns = [
    CONFIG.examNameColumn,
    CONFIG.attendeesColumn,
    CONFIG.unitsColumn,
    CONFIG.minutesPerUnitColumn,
    CONFIG.eksamensHoldColumn
  ];
  
  columns.forEach(column => {
    const range = sheet.getRange(column + CONFIG.startRow + ':' + column + CONFIG.endRow);
    const backgrounds = range.getBackgrounds();
    const clearedBackgrounds = backgrounds.map(row => row.map(color => color === CONFIG.invalidDataColor ? null : color));
    range.setBackgrounds(clearedBackgrounds);
  });
}



function checkHoldConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const holdInterval = parseInt(sheet.getRange(CONFIG.holdInterval).getValue());
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const holdsSchedule = {};

  clearHoldConflictColors();

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
          sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn()).setBackground(CONFIG.dateConflictColor);
        } else {
          holdsSchedule[hold].push({ start: startDate, end: endDate });
        }
      }
    });
  }
}

function checkAttendeeConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const attendeeInterval = parseInt(sheet.getRange(CONFIG.attendeeInterval).getValue());
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const attendeesToCheck = {};

  clearAttendeeConflictColors();

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
            sheet.getRange(event.row, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground(CONFIG.dateConflictColor);
            sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.attendeesColumn + '1').getColumn()).setBackground(CONFIG.dateConflictColor);
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
function checkAttendeeTypos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get the list of valid teachers from the CONFIG.teachersColumn.
  const teachersRange = sheet.getRange(CONFIG.teachersColumn + '2:' + CONFIG.teachersColumn + sheet.getLastRow());
  const teachersValues = teachersRange.getValues();
  const teachers = teachersValues.map(row => row[0].trim());
  // Define the range for attendee data based on CONFIG.
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  const attendeesValues = attendeesRange.getValues();

  // Clear any previous typo-related colors set by this function.
  clearAttendeeTypoColors();

  // Perform the validation checks...
  attendeesValues.forEach((row, i) => {
    const attendeeCell = row[0].trim();
    if (attendeeCell) {
      const attendees = attendeeCell.split(/,\s*/);
      const invalidAttendees = attendees.filter(attendee => !teachers.includes(attendee));

      if (invalidAttendees.length > 0) {
        // If typos are found, set the invalid data color.
        sheet.getRange(CONFIG.attendeesColumn + (i + CONFIG.startRow)).setBackground(CONFIG.invalidDataColor);
      }
    }
  });
}

function checkHoldTypos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get the list of allowed holds from the CONFIG.tilladteHold column.
  const allowedHoldsRange = sheet.getRange(CONFIG.tilladteHold + '2:' + CONFIG.tilladteHold + sheet.getLastRow());
  const allowedHoldsValues = allowedHoldsRange.getValues();
  const allowedHolds = allowedHoldsValues.map(row => row[0].trim());
  // Define the range for hold data based on CONFIG.
  const holdsRange = sheet.getRange(CONFIG.eksamensHoldColumn + CONFIG.startRow + ':' + CONFIG.eksamensHoldColumn + CONFIG.endRow);
  const holdsValues = holdsRange.getValues();

  // Clear any previous typo-related colors set by this function.
  clearHoldTypoColors();

  // Perform the validation checks...
  holdsValues.forEach((row, i) => {
    const holdCell = row[0].trim();
    if (holdCell) {
      const holds = holdCell.split(/,\s*/);
      const invalidHolds = holds.filter(hold => !allowedHolds.includes(hold));

      if (invalidHolds.length > 0) {
        // If typos are found, set the invalid data color.
        sheet.getRange(CONFIG.eksamensHoldColumn + (i + CONFIG.startRow)).setBackground(CONFIG.invalidDataColor);
      }
    }
  });
}

function checkIfStartIsEarlierThanEnd() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Define the ranges for start and end dates based on CONFIG.
  const startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ':' + CONFIG.startDateColumn + CONFIG.endRow);
  const endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ':' + CONFIG.endDateColumn + CONFIG.endRow);
  
  // Clear any previous date conflict-related colors set by this function.
  clearDateConflictColors();

  const startDateValues = startDateRange.getValues().flat();
  const endDateValues = endDateRange.getValues().flat();

  // Perform the checks...
  startDateValues.forEach((startDateValue, i) => {
    const startDate = startDateValue ? new Date(startDateValue) : null;
    const endDate = endDateValues[i] ? new Date(endDateValues[i]) : null;

    if (startDate && endDate && startDate > endDate) {
      // If the start date is later than the end date, set the conflict color.
      startDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateConflictColor);
      endDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateConflictColor);
    }
  });
}

function checkForNotAllowedDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Define the range for not allowed dates based on CONFIG.
  const notAllowedDatesRange = sheet.getRange(CONFIG.notAllowedDates + '2:' + CONFIG.notAllowedDates + (CONFIG.startRow + 15));
  const notAllowedDatesValues = notAllowedDatesRange.getValues();
  
  // Clear any previous date conflict-related colors set by this function.
  clearDateConflictColors();

  // Flatten the notAllowedDatesValues array and convert each to a date object.
  const notAllowedDates = notAllowedDatesValues.flat().map(date => new Date(date).setHours(0, 0, 0, 0));
  const startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ':' + CONFIG.startDateColumn + CONFIG.endRow);
  const endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ':' + CONFIG.endDateColumn + CONFIG.endRow);
  const startDateValues = startDateRange.getValues().flat();
  const endDateValues = endDateRange.getValues().flat();

  // Perform the checks...
  for (let i = 0; i < startDateValues.length; i++) {
    const startDate = startDateValues[i] ? new Date(startDateValues[i]) : null;
    const endDate = endDateValues[i] ? new Date(endDateValues[i]) : null;

    if (startDate && endDate) {
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);

      // Check for any not allowed dates within the start and end dates.
      const hasConflict = notAllowedDates.some(notAllowedDate => notAllowedDate >= startDate && notAllowedDate <= endDate);
      if (hasConflict) {
        // If a date conflict is found, set the conflict color for the start and end date cells.
        startDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateConflictColor);
        endDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateConflictColor);
      }
    }
  }
}


function onEdit(e) {
    clearInvalidInputColors();
    clearDateConflictColors();
    clearAttendeeConflictColors();
    clearHoldConflictColors();

    // Step 2: Check for any empty or malformed inputs
    checkEmptyCells();
    checkAttendeeTypos();
    checkHoldTypos();

    // Step 3: Validate data consistency
    checkIfStartIsEarlierThanEnd();
    isDateTooLate();
    checkForNotAllowedDates();

    // Step 4: Check for conflicts or overlapping data
    checkAttendeeConflicts();
    checkHoldConflicts();

    // Step 5: Calculate any derived or dependent values
    calculateDuration();
    calculateTotalWeight();

    // Step 6: Apply final formatting (if any)
    // This might include setting specific colors or styles based on the calculated values
}



function isDateTooLate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Fetch column index for start date and end date.
  const startDateColIndex = sheet.getRange(CONFIG.startDateColumn + '1').getColumn();
  const endDateColIndex = sheet.getRange(CONFIG.endDateColumn + '1').getColumn();

  // Get the latest permissible date from a specific cell.
  const latestDate = new Date(sheet.getRange(CONFIG.latestDateCell).getValue());

  // Define the range for end dates based on configuration and fetch values.
  const numRows = CONFIG.endRow - CONFIG.startRow + 1;
  const endDateRange = sheet.getRange(CONFIG.startRow, endDateColIndex, numRows, 1);
  const endDateValues = endDateRange.getValues();

  // Clear any previous date conflict-related colors set for late dates.
  clearDateLateColors();

  // Check each end date to see if it exceeds the latest date.
  endDateValues.forEach((endDateValue, i) => {
    const endDate = new Date(endDateValue[0]);

    if (endDate > latestDate) {
      // If the end date is later than the latest date, set the conflict color.
      endDateRange.getCell(i + 1, 1).setBackground(CONFIG.dateConflictColor);
    } else {
      // Clear the color if previously set.
      endDateRange.getCell(i + 1, 1).setBackground(null);
    }
  });
}



function checkEmptyCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Retrieve column indices for various columns
  const examNameColIndex = sheet.getRange(CONFIG.examNameColumn + '1').getColumn();
  const attendeesColIndex = sheet.getRange(CONFIG.attendeesColumn + '1').getColumn();
  const unitsColIndex = sheet.getRange(CONFIG.unitsColumn + '1').getColumn();
  const minutesPerUnitColIndex = sheet.getRange(CONFIG.minutesPerUnitColumn + '1').getColumn();
  const eksamensHoldIndex = sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn();

  const startRow = CONFIG.startRow;
  const endRow = CONFIG.endRow;
  const numRows = endRow - startRow + 1;
  
  // Clear previous invalid data colors
  clearInvalidInputColors();

  // Loop through each row to check for missing inputs
  for (let i = 0; i < numRows; i++) {
    const row = sheet.getRange(startRow + i, 1, 1, sheet.getMaxColumns()).getValues()[0];
    
    if (toStringAndTrim(row[attendeesColIndex - 1])) { // Check if there's data in the attendees column
      const examNameColor = toStringAndTrim(row[examNameColIndex - 1]) ? null : CONFIG.invalidDataColor; 
      const unitsColor = toStringAndTrim(row[unitsColIndex - 1]) ? null : CONFIG.invalidDataColor; 
      const minutesPerUnitColor = toStringAndTrim(row[minutesPerUnitColIndex - 1]) ? null : CONFIG.invalidDataColor; 
      const eksamensHoldColor = toStringAndTrim(row[eksamensHoldIndex - 1]) ? null : CONFIG.invalidDataColor; 

      sheet.getRange(startRow + i, examNameColIndex).setBackground(examNameColor);
      sheet.getRange(startRow + i, unitsColIndex).setBackground(unitsColor);
      sheet.getRange(startRow + i, minutesPerUnitColIndex).setBackground(minutesPerUnitColor);
      sheet.getRange(startRow + i, eksamensHoldIndex).setBackground(eksamensHoldColor);
    } else { // Clear colors if no attendee data
      sheet.getRange(startRow + i, examNameColIndex).setBackground(null);
      sheet.getRange(startRow + i, unitsColIndex).setBackground(null);
      sheet.getRange(startRow + i, minutesPerUnitColIndex).setBackground(null);
      sheet.getRange(startRow + i, eksamensHoldIndex).setBackground(null);
    }
  }
}

function toStringAndTrim(value) {
  // Convert any value to string and trim it
  return value === null || value === undefined ? '' : String(value).trim();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🪄EASV Koordinator Powertools🪄')
    .addItem('Udfyld resten af datoerne for mig⚡', 'generateDates')
    .addItem('Lav skemaer ud fra Moodle grupper (CSV i kolonne M)✨', 'createStudentSchedule')
    .addToUi();
}
function generateDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const earliestDate = new Date(sheet.getRange(CONFIG.earliestDateCell).getValue());
  earliestDate.setHours(0, 0, 0, 0); 
  const maksTid = parseFloat(sheet.getRange(CONFIG.maksTid).getValue());
  const attendeesSchedule = {};
  const interval = sheet.getRange(CONFIG.interval).getValue();
  const notAllowedDatesRange = sheet.getRange(CONFIG.notAllowedDates + "2:" + CONFIG.notAllowedDates + (CONFIG.startRow + 15));
  const notAllowedDates = notAllowedDatesRange.getValues().flat().map(date => new Date(date).setHours(0, 0, 0, 0));

    for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const attendees = row[sheet.getRange(CONFIG.attendeesColumn + '1').getColumn() - 1].trim().split(/,\s*/);
    const startDateCell = sheet.getRange(CONFIG.startDateColumn + (i + CONFIG.startRow));
    const endDateCell = sheet.getRange(CONFIG.endDateColumn + (i + CONFIG.startRow));
    const startDate = startDateCell.getValue();
    const endDate = endDateCell.getValue();

    if (startDate && endDate) {
      const startDateTime = new Date(startDate);
      const endDateTime = new Date(endDate);
      startDateTime.setHours(0, 0, 0, 0);
      endDateTime.setHours(23, 59, 59, 999);
      attendees.forEach(attendee => {
        if (!attendeesSchedule[attendee]) {
          attendeesSchedule[attendee] = [];
        }
        attendeesSchedule[attendee].push({ start: startDateTime, end: endDateTime });
      });
    }
  }

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const attendees = row[sheet.getRange(CONFIG.attendeesColumn + '1').getColumn() - 1].trim().split(/,\s*/);
    let totalTime = row[sheet.getRange(CONFIG.totalTidColumn + '1').getColumn() - 1];
    if (totalTime <= maksTid) totalTime = maksTid;
    const startDateCell = sheet.getRange(CONFIG.startDateColumn + (i + CONFIG.startRow));
    const endDateCell = sheet.getRange(CONFIG.endDateColumn + (i + CONFIG.startRow));

    if (!startDateCell.getValue() && !endDateCell.getValue() && totalTime && attendees.length > 0 && JSON.stringify(attendees).trim().length > 4) {
       const totalDays = Math.ceil(totalTime / maksTid);
      let startDate = new Date(earliestDate);
      let endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + totalDays - 1);
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);

      let attempts = 0;
      const maxAttempts = 100; 

      while (attempts < maxAttempts) {
        
        if (notAllowedDates.some(date => date >= startDate && date <= endDate)) {
          startDate.setDate(startDate.getDate() + 1);
          endDate.setDate(endDate.getDate() + 1);
        } else if (attendees.some(attendee => attendeesSchedule[attendee]?.some(event => event.start < endDate && startDate < event.end))) {
          startDate.setDate(startDate.getDate() + 1 + interval);
          endDate.setDate(endDate.getDate() + 1 + interval);
        } else {
          break;
        }
        attempts++;
      }

      if (attempts < maxAttempts) {
        const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        const formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        startDateCell.setValue(formattedStartDate);
        endDateCell.setValue(formattedEndDate);
        attendees.forEach(attendee => {
          if (!attendeesSchedule[attendee]) {
            attendeesSchedule[attendee] = [];
          }
          attendeesSchedule[attendee].push({ start: startDate, end: endDate });
        });
      }
    }
  }  
  isDateTooLate();
  checkAttendeeConflicts();
  checkHoldConflicts();

}



function calculateTotalWeight() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  const totalTidRange = sheet.getRange(CONFIG.totalTidColumn + CONFIG.startRow + ':' + CONFIG.totalTidColumn + CONFIG.endRow);
  const teachersRange = sheet.getRange(CONFIG.teachersColumn + '2:' + CONFIG.teachersColumn + sheet.getLastRow());
  const totalBelastningPrInitial = sheet.getRange(CONFIG.totalWeightPerEmployeeColumn + '2:' + CONFIG.totalWeightPerEmployeeColumn + sheet.getLastRow());
  const attendeesValues = attendeesRange.getValues();
  const totalTidValues = totalTidRange.getValues();
  const teachersValues = teachersRange.getValues().map(row => row[0].trim());

  let totalTimes = {};

  attendeesValues.forEach((row, i) => {
    const attendees = row[0].trim().split(/,\s*/);
    const totalTime = parseFloat(totalTidValues[i][0]); 

    attendees.forEach(attendee => {
      if (!totalTimes[attendee]) {
        totalTimes[attendee] = 0;
      }
      totalTimes[attendee] += isNaN(totalTime) ? 0 : totalTime;
    });
  });

  totalBelastningPrInitial.clearContent();
  teachersValues.forEach((teacherInitials, i) => {
    if (totalTimes[teacherInitials] !== undefined && teacherInitials.length > 0) {
      const weightCell = sheet.getRange(i + 2, sheet.getRange(CONFIG.totalWeightPerEmployeeColumn + '1').getColumn());
      weightCell.setValue(totalTimes[teacherInitials]);
    }
  });
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
    const totalTidCell = sheet.getRange(CONFIG.totalTidColumn + (i + CONFIG.startRow));

    if (numUnits && minutesPerUnit && !isNaN(numUnits) && !isNaN(minutesPerUnit) && numUnits > 0 && minutesPerUnit > 0) {
      const totalTimeInHours = (numUnits * minutesPerUnit) / 60;
      
      if (totalTimeInHours > 0) {
        totalTidCell.setValue(totalTimeInHours);
        totalTidCell.setNumberFormat('#,##0.00');
      } else {
        totalTidCell.setValue("");
      }

      if (totalTimeInHours > 50) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Advarsel om mulig fejl-indtastning: Beregnet tid er over 50 timer, tjek venligst minutter per eksamen og holdstørrelse");
      }
    } else {
      totalTidCell.setValue("");
    }
  }
}


function createStudentSchedule() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getActiveSheet();
  const dataRange = mainSheet.getDataRange();
  const dataValues = dataRange.getValues();
  
  const examNameColIndex = getColumnIndex(CONFIG.examNameColumn);
  const minutesPerUnitColIndex = getColumnIndex(CONFIG.minutesPerUnitColumn);
  const csvDataColIndex = getColumnIndex(CONFIG.csvDataColumn);

  for (let i = CONFIG.startRow - 1; i < CONFIG.endRow; i++) {
    const examName = dataValues[i][examNameColIndex - 1];
    const csvData = dataValues[i][csvDataColIndex - 1];
    const minutesPerUnit = dataValues[i][minutesPerUnitColIndex - 1];

    if (csvData && examName) {
      const groups = Utilities.parseCsv(csvData);
      let scheduleSheet = spreadsheet.getSheetByName(examName);
      if (!scheduleSheet) {
        scheduleSheet = spreadsheet.insertSheet(examName);
      } else {
        scheduleSheet.clear(); 
      }

      scheduleSheet.appendRow(['Student (full name)', 'Group name', 'Starting time']);

      let startTime = new Date();
      startTime.setHours(9, 0, 0, 0); 

      groups.forEach(group => {
        const groupName = group[1]; 
        for (let memberIndex = 8; memberIndex < group.length; memberIndex += 4) {
          if (group[memberIndex]) { 
            const fullName = group[memberIndex + 2] + ' ' + group[memberIndex + 3]; 
            scheduleSheet.appendRow([
              fullName,
              groupName,
              Utilities.formatDate(startTime, spreadsheet.getSpreadsheetTimeZone(), 'HH:mm')
            ]);

            startTime = new Date(startTime.getTime() + minutesPerUnit * 60000);
          }
        }
      });
    }
  }
}

function getColumnIndex(columnLetter) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnLetter + '1').getColumn();
}