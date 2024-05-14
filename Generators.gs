function createStudentSchedule() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getActiveSheet();
  const dataRange = mainSheet.getDataRange();
  const dataValues = dataRange.getValues();
  
  const examNameColIndex = CONFIG.examNameColumn.charCodeAt(0) - 'A'.charCodeAt(0);
  const minutesPerUnitColIndex = CONFIG.minutesPerUnitColumn.charCodeAt(0) - 'A'.charCodeAt(0);
  const csvDataColIndex = CONFIG.csvDataColumn.charCodeAt(0) - 'A'.charCodeAt(0);

  for (let i = CONFIG.startRow - 1; i < CONFIG.endRow; i++) {
    const examName = dataValues[i][examNameColIndex];
    const csvData = dataValues[i][csvDataColIndex];
    const minutesPerUnit = parseInt(dataValues[i][minutesPerUnitColIndex]);

    if (csvData && examName) {
      const groups = Utilities.parseCsv(csvData);
      
      let scheduleSheet = spreadsheet.getSheetByName(examName);
      if (!scheduleSheet) {
        scheduleSheet = spreadsheet.insertSheet(examName);
      } else {
        scheduleSheet.clear();
      }

      scheduleSheet.appendRow(['Student Identifier', 'Group name', 'Starting time']);

      let startTime = new Date();
      startTime.setHours(9, 0, 0, 0);

      for (let j = 1; j < groups.length; j++) {
        const group = groups[j];
        if (group.length > 8 && group[1]) { 
          const groupName = group[1];
          for (let memberIndex = 8; memberIndex + 4 < group.length; memberIndex += 5) {
            const username = group[memberIndex];
            const idNumber = group[memberIndex + 1];
            const firstName = group[memberIndex + 2];
            const lastName = group[memberIndex + 3];
            const email = group[memberIndex + 4]; 

            const identifier = firstName && lastName ? `${firstName} ${lastName}` : (username || email || 'Unknown Identifier');

            if (identifier && (firstName || lastName || username || email)) {
              scheduleSheet.appendRow([
                identifier,
                groupName,
                Utilities.formatDate(startTime, spreadsheet.getSpreadsheetTimeZone(), 'HH:mm')
              ]);

              startTime = new Date(startTime.getTime() + minutesPerUnit * 60000);
            }
          }
        }
      }
    }
  }
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


}