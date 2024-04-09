function checkAttendeeConflicts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var teachersColumn = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  var teachers = teachersColumn.map(function(row) { return row[0].trim(); });

  // Clear the background for the entire data range
  range.setBackground(null);

  var attendeesToCheck = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var startDate = new Date(row[5]).setHours(0, 0, 0, 0); // Changed index to 5
    var endDate = new Date(row[6]).setHours(23, 59, 59, 999); // Changed index to 6
    var attendees = row[4].split(','); // Changed index to 4

    var allAttendeesValid = attendees.every(function(attendee) {
      return teachers.includes(attendee.trim());
    });

    if (!allAttendeesValid) {
      sheet.getRange(i + 1, 5).setBackground('red'); // Changed index to 5
    }

    for (var j = 0; j < attendees.length; j++) {
      var attendee = attendees[j].trim();

      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + 1 }];
      } else {
        var conflicts = attendeesToCheck[attendee].some(function(event) {
          var hasConflict = (startDate <= event.end && endDate >= event.start);
          if (hasConflict) {
            sheet.getRange(event.row, 5).setBackground('orange'); // Changed index to 5
          }
          return hasConflict;
        });

        if (conflicts) {
          sheet.getRange(i + 1, 5).setBackground('orange'); // Changed index to 5
        }
        attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + 1 });
      }
    }
  }
}

function onEdit(e) {
  var attendeesColumnIndex = 5; // Examinators
  var startDateColumnIndex = 6; // Start Date
  var endDateColumnIndex = 7; // End Date

  if ([attendeesColumnIndex, startDateColumnIndex, endDateColumnIndex].includes(e.range.columnStart)) {
    checkAttendeeConflicts();
  }
}
