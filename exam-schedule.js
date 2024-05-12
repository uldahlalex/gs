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

  startRow: 2,
  endRow: 100,
  conflictColor: "#e391e3", 


  maksTid: 'O2',
  interval: 'O3',
  earliestDateCell: 'O4',
  latestDateCell: 'O8',
    holdInterval: 'O9',

};



function checkHoldConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.startRow, 1, sheet.getLastRow() - CONFIG.startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const holdsToCheck = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const startDate = new Date(row[sheet.getRange(CONFIG.startDateColumn + '1').getColumn() - 1]);
    const endDate = new Date(row[sheet.getRange(CONFIG.endDateColumn + '1').getColumn() - 1]);
    const hold = row[sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn() - 1].trim();

    if (hold) { // Only proceed if there's a group assigned
      if (!holdsToCheck[hold]) {
        holdsToCheck[hold] = [{ start: startDate, end: endDate, row: i + CONFIG.startRow }];
      } else {
        const conflicts = holdsToCheck[hold].some(event => {
          const hasConflict = (startDate < event.end && endDate > event.start);
          if (hasConflict) {
            // Color the conflicting hold cell in 'orange'
            sheet.getRange(event.row, sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn()).setBackground(CONFIG.conflictColor);
            sheet.getRange(i + CONFIG.startRow, sheet.getRange(CONFIG.eksamensHoldColumn + '1').getColumn()).setBackground(CONFIG.conflictColor);
          }
          return hasConflict;
        });

        if (!conflicts) {
          holdsToCheck[hold].push({ start: startDate, end: endDate, row: i + CONFIG.startRow });
        }
      }
    }
  }
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

function checkAttendeeTypos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const teachersRange = sheet.getRange(CONFIG.teachersColumn + '2:' + CONFIG.teachersColumn + sheet.getLastRow());
  const teachersValues = teachersRange.getValues();
  const teachers = teachersValues.map(row => row[0].trim());
  const attendeesRange = sheet.getRange(CONFIG.attendeesColumn + CONFIG.startRow + ':' + CONFIG.attendeesColumn + CONFIG.endRow);
  const attendeesValues = attendeesRange.getValues();

  attendeesRange.setBackground(null);

  attendeesValues.forEach((row, i) => {
    const attendeeCellContent = row[0].trim();
    if (attendeeCellContent) { 
      const attendees = attendeeCellContent.split(/,\s*/);
      const allAttendeesValid = attendees.every(attendee => teachers.includes(attendee.trim()));
      if (!allAttendeesValid) {
        sheet.getRange(CONFIG.attendeesColumn + (i + CONFIG.startRow)).setBackground('red');
      }
    }
  });
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
      colorRedIfLackingInputs();
  if ([CONFIG.attendeesColumn, CONFIG.tilladteHold, CONFIG.eksamensHoldColumn, CONFIG.teachersColumn, CONFIG.startDateColumn, CONFIG.endDateColumn, CONFIG.notAllowedDates].includes(e.range.getA1Notation().charAt(0))) {
    checkAttendeeTypos();
    checkAttendeeConflicts();
    checkDateConflictsAndColorCells();
    checkHoldConflicts();
    checkHoldTypos();
  }

  if([CONFIG.startDateColumn, CONFIG.eksamensHoldColumn, CONFIG.endDateColumn].includes(e.range.getA1Notation().charAt(0))) {
    dateValidation();

  }

    calculateDuration();
    calculateTotalWeight();


}


function checkHoldTypos() {
  //just like attendee conflicts
}

function dateValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const startDateColIndex = sheet.getRange(CONFIG.startDateColumn + '1').getColumn();

  const latestDate = new Date(sheet.getRange(CONFIG.latestDateCell).getValue());
  
  const numRows = CONFIG.endRow - CONFIG.startRow + 1;
  const dateRange = sheet.getRange(CONFIG.startRow, startDateColIndex, numRows, 2); 
  const dateValues = dateRange.getValues();

  for (let i = 0; i < numRows; i++) {
    const startDate = new Date(dateValues[i][0]);
    const endDate = new Date(dateValues[i][1]); 

    if (startDate > endDate || endDate > latestDate) {
      dateRange.getCell(i + 1, 1).setBackground(CONFIG.conflictColor); 
      dateRange.getCell(i + 1, 2).setBackground(CONFIG.conflictColor);
    } else {
      dateRange.getCell(i + 1, 1).setBackground('#FFFFFF'); 
      dateRange.getCell(i + 1, 2).setBackground('#FFFFFF'); 
    }
  }
}
function colorRedIfLackingInputs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
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
    
    if (String(row[attendeesColIndex - 1]).trim()) {
      const examNameColor = String(row[examNameColIndex - 1]).trim() ? '#FFFFFF' : '#FF0000'; 
      const unitsColor = String(row[unitsColIndex - 1]).trim() ? '#FFFFFF' : '#FF0000'; 
      const minutesPerUnitColor = String(row[minutesPerUnitColIndex - 1]).trim() ? '#FFFFFF' :
       '#FF0000'; 
             const eksamensHoldColor = String(row[eksamensHoldIndex - 1]).trim() ? '#FFFFFF' :
       '#FF0000'; 

      range.getCell(i + 1, examNameColIndex).setBackground(examNameColor);
      range.getCell(i + 1, unitsColIndex).setBackground(unitsColor);
      range.getCell(i + 1, minutesPerUnitColIndex).setBackground(minutesPerUnitColor);
            range.getCell(i + 1, eksamensHoldIndex).setBackground(eksamensHoldColor);

    } else {
      range.getCell(i + 1, examNameColIndex).setBackground('#FFFFFF');
      range.getCell(i + 1, unitsColIndex).setBackground('#FFFFFF');
      range.getCell(i + 1, minutesPerUnitColIndex).setBackground('#FFFFFF');
            range.getCell(i + 1, eksamensHoldIndex).setBackground('#FFFFFF');

    }
  }
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
  checkAttendeeConflicts();
  dateValidation();
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