function checkAttendeeConflicts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var teachersColumn = sheet.getRange(2, 8, sheet.getLastRow() - 1).getValues();
  var teachers = teachersColumn.map(function(row) { return row[0].trim(); });

  range.setBackground(null);

  var attendeesToCheck = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var startDate = new Date(row[2]).setHours(0, 0, 0, 0);
    var endDate = new Date(row[3]).setHours(23, 59, 59, 999);
    var attendees = row[1].split(',');

    var allAttendeesValid = attendees.every(function(attendee) {
      return teachers.includes(attendee.trim());
    });

    if (!allAttendeesValid) {
      sheet.getRange(i + 1, 2).setBackground('red');
    }
    for (var j = 0; j < attendees.length; j++) {
      var attendee = attendees[j].trim();

      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + 1 }];
      } else {
        var conflicts = attendeesToCheck[attendee].some(function(event) {
          var hasConflict = (startDate <= event.end && endDate >= event.start);
          if (hasConflict) {
            sheet.getRange(event.row, 2).setBackground('orange');
          }
          return hasConflict;
        });

        if (conflicts) {
          sheet.getRange(i + 1, 2).setBackground('orange');
        }
        attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + 1 });
      }
    }
  }
}

function onEdit(e) {
  var attendeesColumnIndex = 2;
  var startDateColumnIndex = 3;
  var endDateColumnIndex = 4;

  if ([attendeesColumnIndex, startDateColumnIndex, endDateColumnIndex].includes(e.range.columnStart)) {
    checkAttendeeConflicts();
  }
}
