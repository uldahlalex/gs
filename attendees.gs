function checkAttendeeConflicts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Get the list of teachers from column H (index 7)
  var teachersColumn = sheet.getRange(2, 8, sheet.getLastRow() - 1).getValues();
  var teachers = teachersColumn.map(function(row) { return row[0].trim(); });

  // Clear any previous formatting
  range.setBackground(null);

  var attendeesToCheck = {};

  for (var i = 1; i < values.length; i++) { // Start from 1 to skip header row
    var row = values[i];
    var startDate = new Date(row[2]).setHours(0, 0, 0, 0); // Start Date at midnight
    var endDate = new Date(row[3]).setHours(23, 59, 59, 999); // End Date just before midnight
    var attendees = row[1].split(','); // Change delimiter if necessary

    // Check if all attendees are in the list of teachers
    var allAttendeesValid = attendees.every(function(attendee) {
      return teachers.includes(attendee.trim());
    });

    // Invalid attendees are highlighted in red
    if (!allAttendeesValid) {
      sheet.getRange(i + 1, 2).setBackground('red');
      continue; // Skip further processing for invalid rows
    }

    // Process valid attendees for conflicts
    for (var j = 0; j < attendees.length; j++) {
      var attendee = attendees[j].trim();

      if (!attendeesToCheck[attendee]) {
        attendeesToCheck[attendee] = [{ start: startDate, end: endDate, row: i + 1 }];
      } else {
        var conflicts = attendeesToCheck[attendee].some(function(event) {
          var hasConflict = (startDate <= event.end && endDate >= event.start);
          if (hasConflict) {
            // Highlight the conflicting event's row as well
            sheet.getRange(event.row, 2).setBackground('orange');
          }
          return hasConflict;
        });

        if (conflicts) {
          // Highlights the Attendees cell for date conflicts
          sheet.getRange(i + 1, 2).setBackground('orange');
        }
        attendeesToCheck[attendee].push({ start: startDate, end: endDate, row: i + 1 });
      }
    }
  }
}

function onEdit(e) {
  // Define the columns you want to watch for edits
  var attendeesColumnIndex = 2; // Column B
  var startDateColumnIndex = 3; // Column C
  var endDateColumnIndex = 4; // Column D

  // Check if the edit was made in one of the specified columns
  if ([attendeesColumnIndex, startDateColumnIndex, endDateColumnIndex].includes(e.range.columnStart)) {
    // Call your conflict check function if the edit is in one of the relevant columns
    checkAttendeeConflicts();
  }
}
