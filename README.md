

## Contents:

- exam-schedule.js: This is actually the .gs script attached as an Apps Script to Google Sheets (saved in .js to get syntax highlighting)
- testSpreadsheet: Textual data for the testing spreadsheet (doens't include formatting)



## Roadmap:
- Correction of timetable generation for individual exams (needs slightly more data to be useful and the header row should be skipped)
    - Should respect lunch break
    - Should span several days if required (start-end dates implementation instead of "just" starting at 9 o'clock at start day and running endlessly)
    - Should create new sheet every time for every exam the generator is invoked (no auto overwrite / delete)
    - Should always add identifier for student if no name can be found
- Checking Moodle "Group self select" preferences doesn't break the CSV output required for the generate timetable function (if this is the case, maybe a custom Moodle plugin will help?)

## Live test:
Simply go to File -> Make a Copy to get the spreadsheet with the Apps Script from here:
https://docs.google.com/spreadsheets/d/1HGoarGPOfJubhPFkO7slvUpsQNZFTtgk2x3HrsSHxZU/edit#gid=0