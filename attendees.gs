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

    if (!allAttendeesValid) {
      sheet.getRange(i + 1, 2).setBackground('red'); // Invalid attendees are highlighted
    } else {
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
}
