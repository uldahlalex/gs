function calculateTotalWeight(sheet) {
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

function calculateDuration(sheet) {
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
