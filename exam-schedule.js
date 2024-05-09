const CONFIG = {
  teachersColumn: 'A',
  totalWeightPerEmployeeColumn: 'B',
  notAllowedDates: 'C',
  attendeesColumn: 'E',
  startDateColumn: 'F',
  endDateColumn: 'G',
  unitsColumn: 'H',
  minutesPerUnitColumn: 'I',
  totalTidColumn: 'J',
  startRow: 2,
  endRow: 100,
  conflictColor: "#e391e3", 
  dataCheckColumns: ['H', 'I'], //til beregning af maks tid
  nonEditableColumns: ['J'],
     maksTid: 'L2',
     interval: 'L3',
  earliestDateCell: 'L4',

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
  var notAllowedDatesRange = sheet.getRange(CONFIG.notAllowedDates + "2:" + CONFIG.notAllowedDates + (CONFIG.startRow + 15));
  var notAllowedDates = notAllowedDatesRange.getValues().flat();
  var startDateRange = sheet.getRange(CONFIG.startDateColumn + CONFIG.startRow + ":" + CONFIG.startDateColumn);
  var startDateValues = startDateRange.getValues().flat();
  var endDateRange = sheet.getRange(CONFIG.endDateColumn + CONFIG.startRow + ":" + CONFIG.endDateColumn);
  var endDateValues = endDateRange.getValues().flat();

  notAllowedDatesRange.setBackground(null);
  startDateRange.setBackground(null);
  endDateRange.setBackground(null);

  for (var i = 0; i < startDateValues.length; i++) {
    var startDate = startDateValues[i];
    var endDate = endDateValues[i];
    if (startDate && endDate) { 
      startDate = new Date(startDate);
      endDate = new Date(endDate);
      for (var j = 0; j < notAllowedDates.length; j++) {
        var notAllowedDate = new Date(notAllowedDates[j]);
        if (notAllowedDate >= startDate && notAllowedDate <= endDate) {
          notAllowedDatesRange.getCell(j + 1, 1).setBackground(CONFIG.conflictColor);
          startDateRange.getCell(i + 1, 1).setBackground(CONFIG.conflictColor);
          endDateRange.getCell(i + 1, 1).setBackground(CONFIG.conflictColor);
        }
      }
    }
  }
}
function onEdit(e) {
  var range = e.range;
  
  if ([CONFIG.attendeesColumn, CONFIG.startDateColumn, CONFIG.endDateColumn, CONFIG.notAllowedDates].includes(e.range.getA1Notation().charAt(0))) {
    checkAttendeeTypos();
    checkAttendeeConflicts();
    checkDateConflictsAndColorCells();
  }

    calculateDuration();
    calculateTotalWeight();
    colorRedIfLackingInputs();

}

function colorRedIfLackingInputs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const startRow = 2; 
  const lastRow = sheet.getLastRow();
  
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + startRow + ':'+CONFIG.attendeesColumn + lastRow);
  const attendeesValues = attendeesRange.getValues();
  
  for (let i = 0; i < attendeesValues.length; i++) {
    if (attendeesValues[i][0]) {  
      const hCell = sheet.getRange(startRow + i, 8);  
      const iCell = sheet.getRange(startRow + i, 9);  
      
      hCell.setBackground(hCell.getValue() ? '#FFFFFF' : '#FF0000');
      iCell.setBackground(iCell.getValue() ? '#FFFFFF' : '#FF0000');
    }
  }
}
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  sheet.toast('Dette ark er endnu ikke klar til at blive kopieret og anvendt', 'Regneark under redigering', 0); // message, title, timeout in seconds
  SpreadsheetApp.getUi()
    .createMenu('🪄EASV Koordinator Powertools🪄')
    .addItem('Udfyld resten af datoerne for mig⚡', 'generateDates')
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
  checkAttendeeConflicts();
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
    if (numUnits && minutesPerUnit && !isNaN(numUnits) && !isNaN(minutesPerUnit)) {
      const totalTimeInHours = (numUnits * minutesPerUnit) / 60;
      const totalTidCell = sheet.getRange(CONFIG.totalTidColumn + (i + CONFIG.startRow));
      totalTidCell.setValue(totalTimeInHours);
      totalTidCell.setNumberFormat('#,##0.00');
      if (totalTimeInHours > 50) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Advarsel om mulig fejl-indtastning: Beregnet tid er over 50 timer, tjek venligst minutter per eksamen og holdstørrelse");
      }
    } else {
          const totalTidCell = sheet.getRange(CONFIG.totalTidColumn + (i + CONFIG.startRow));
                totalTidCell.setValue(0);
    }
  }
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