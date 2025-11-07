# late-sheet

Google Apps Script project for dxe.io/late Late Sheet

## Deployment

Copy and paste into Apps Script editor. The script is contained in the
spreadsheet.

Create a trigger that runs hourly and runs "processLateSheetIfNotModifiedRecently".

## Considerations

Do not use Google Sheets "tables" feature because Google Apps Script does not support it
and cannot find the validation rule.
