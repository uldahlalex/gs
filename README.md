
## Roadmap:
- Correction of timetable generation for individual exams
    - Should respect lunch break
    - Should span several days if required (start-end dates implementation instead of "just" starting at 9 o'clock at start day and running endlessly)
    - Should create new sheet every time for every exam the generator is invoked (no auto overwrite / delete)
    - Should always add identifier for student if no name can be found
- Checking Moodle "Group self select" preferences doesn't break the CSV output required for the generate timetable function (if this is the case, maybe a custom Moodle plugin will help?)
- Documentation and help resouces
    - Videomaterial once the new versoin is ready (+ link to this in spreadsheet)
    - Keymap / column identifiers
    - Written guide inside the spreadsheet? or outside it with link to this in the spreadsheet
    - 

## Live test:
Simply go to File -> Make a Copy to get the spreadsheet with the Apps Script from here:
https://docs.google.com/spreadsheets/d/1HGoarGPOfJubhPFkO7slvUpsQNZFTtgk2x3HrsSHxZU/edit#gid=0