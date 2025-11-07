const spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1j_3it9c1WMn9z_1r4qvENPPcd_Sl3jDImR1GLZwN0Lo";

const columns = {
    "name": "Name",
    "date": "Date",
    "violation": "Violation",
    "howLate": "How Late?",
    "notes": "Notes",
    "emailSent": "Email sent",
    "coachNotified": "Coach notified",
}

/**
 * Check if the spreadsheet was modified within the last 10 minutes
 * @param {Spreadsheet} spreadsheet - The spreadsheet to check
 * @returns {boolean} True if spreadsheet was modified within last 10 minutes, false otherwise
 */
function wasSheetModifiedRecently(spreadsheet) {
  try {
    const file = DriveApp.getFileById(spreadsheet.getId());
    const lastModified = file.getLastUpdated();
    
    const now = new Date();
    const tenMinutesAgo = new Date(now.getTime() - 10 * 60 * 1000);
    
    if (lastModified > tenMinutesAgo) {
      Logger.log(`Sheet was last modified at ${lastModified}, which is within the last 10 minutes.`);
      return true;
    }
    
    return false;
  } catch (e) {
    Logger.log(`Error checking sheet modification time: ${e.toString()}`);
    // If we can't check, don't skip processing (process normally)
    return false;
  }
}

/**
 * Main entry point
 */
function processLateSheetIfNotModifiedRecently() {
  const spreadsheet = getSpreadsheetForThisYear();
  if (!spreadsheet) {
    return;
  }
  
  // Skip processing if sheet was modified recently
  if (wasSheetModifiedRecently(spreadsheet)) {
    Logger.log(`Skipping processing of sheet since it was modified in the last 10 minutes`);
    return;
  }
  
  processLateSheet(spreadsheet);
}

function processLateSheetNow() {
   const spreadsheet = getSpreadsheetForThisYear();
  if (!spreadsheet) {
    return;
  } 
  processLateSheet(spreadsheet);
}

function processLateSheet(spreadsheet) {
  const currentYear = new Date().getFullYear().toString();
  if (currentYear === "2025") {
    Logger.log(`Don't start until 2026`);
    return;
  }
  const sheet = spreadsheet.getSheetByName(currentYear);
  if (!sheet) {
    Logger.log(`Sheet named "${currentYear}" not found`);
    return;
  }
  
  processSheet(spreadsheet, sheet);
}

/**
 * Get the spreadsheet for the current year
 * @returns {Spreadsheet|null} The spreadsheet or null if not found
 */
function getSpreadsheetForThisYear() {
  return SpreadsheetApp.openByUrl(spreadsheetUrl);
}

/**
 * Read column headers and map them to indices, and get validation rule for name column
 * @param {Sheet} sheet - The sheet to read from
 * @returns {Object|null} Object mapping column keys to indices, values, and nameValidation
 */
function readColumnData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 2) {
    Logger.log("No data rows found");
    return null;
  }
  
  // Read column names from the first row
  const headerRow = values[0];
  const columnIndices = {};
  
  // Map column headers to their indices
  headerRow.forEach((header, index) => {
    Object.keys(columns).forEach(key => {
      if (columns[key] === header) {
        columnIndices[key] = index;
      }
    });
  });
  
  // Verify all defined columns exist
  const missingColumns = [];
  Object.keys(columns).forEach(key => {
    if (columnIndices[key] === undefined) {
      missingColumns.push(columns[key]);
    }
  });
  
  if (missingColumns.length > 0) {
    Logger.log(`Missing columns: ${missingColumns.join(', ')}`);
    return null;
  }
  
  // Get validation rule for the name column from a data cell (row 2, first data row)
  let nameValidation = null;
  const nameColumnIndex = columnIndices.name;
  if (nameColumnIndex !== undefined) {
    try {
      // Try to get validation from the first data row (row 2, since row 1 is headers)
      const nameCell = sheet.getRange(2, nameColumnIndex + 1);
      nameValidation = nameCell.getDataValidation();
      
      if (nameValidation) {
        Logger.log(`Found validation rule for "${columns.name}" column`);
      } else {
        Logger.log(`No validation rule found for "${columns.name}" column`);
      }
    } catch (e) {
      Logger.log(`Error getting validation rule: ${e.toString()}`);
    }
  }
  
  return {
    columnIndices: columnIndices,
    values: values,
    nameValidation: nameValidation
  };
}

/**
 * Check if the name value is valid according to the validation rule
 * @param {DataValidation|null} validation - The validation rule (can be null)
 * @param {string} nameValue - The name value to validate
 * @param {Spreadsheet} spreadsheet - The spreadsheet (needed for range validation)
 * @returns {boolean} True if the name is valid, false otherwise
 */
function isValidName(validation, nameValue, spreadsheet) {
  if (!nameValue || nameValue.toString().trim() === '') {
    return false;
  }
  
  if (!validation) {
    // No validation rule, consider it valid
    return true;
  }
  
  try {
    const criteriaType = validation.getCriteriaType();
    const criteriaValues = validation.getCriteriaValues();
    
    Logger.log(`Validating name "${nameValue}" against criteria type: ${criteriaType}`);
    
    // Check if it's a dropdown validation (VALUE_IN_LIST or VALUE_IN_RANGE)
    if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      if (criteriaValues && criteriaValues.length > 0 && criteriaValues[0]) {
        const allowedValues = criteriaValues[0];
        const nameStr = nameValue.toString().trim();
        const isValid = allowedValues.some(val => val.toString().trim() === nameStr);
        Logger.log(`Validation result for "${nameValue}": ${isValid ? 'VALID' : 'INVALID'} (checked against ${allowedValues.length} allowed values)`);
        return isValid;
      }
      Logger.log(`VALUE_IN_LIST validation found but no criteria values available`);
      // If validation exists but no values, consider valid to avoid blocking
      return true;
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      if (criteriaValues && criteriaValues.length > 0 && criteriaValues[0]) {
        try {
          // For range validation, get the range and check if value exists in it
          const rangeA1 = criteriaValues[0].toString();
          const range = spreadsheet.getRange(rangeA1);
          const rangeValues = range.getValues().flat();
          const nameStr = nameValue.toString().trim();
          const isValid = rangeValues.some(val => val.toString().trim() === nameStr);
          Logger.log(`Validation result for "${nameValue}": ${isValid ? 'VALID' : 'INVALID'} (checked against range ${rangeA1})`);
          return isValid;
        } catch (e) {
          Logger.log(`Error checking range validation: ${e.toString()}`);
          // If we can't check the range, consider it valid to avoid blocking entries
          return true;
        }
      }
      Logger.log(`VALUE_IN_RANGE validation found but no criteria values available`);
      // If validation exists but no values, consider valid to avoid blocking
      return true;
    } else {
      Logger.log(`Unknown validation criteria type: ${criteriaType}, considering valid`);
      // For unknown types, consider valid to avoid blocking entries
      return true;
    }
  } catch (e) {
    Logger.log(`Error validating name: ${e.toString()}`);
    // If we can't validate due to error, consider valid to avoid blocking entries
    return true;
  }
  
  // If we reach here, validation wasn't properly handled, consider valid
  return true;
}

/**
 * Process the sheet and send emails for unprocessed rows
 * @param {Spreadsheet} spreadsheet - The spreadsheet
 * @param {Sheet} sheet - The sheet to process
 */
function processSheet(spreadsheet, sheet) {
  if (!sheet) {
    return;
  }
  
  // Read column data
  const columnData = readColumnData(sheet);
  if (!columnData) {
    return;
  }
  
  const { columnIndices, values, nameValidation } = columnData;
  
  // Find unprocessed rows where emailSent is blank
  const unprocessedRows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const emailSent = row[columnIndices.emailSent];
    
    // Check if emailSent is blank
    if (!emailSent || emailSent.toString().trim() === '') {
      // Parse the row into an object
      const rowData = {
        rowNumber: i + 1, // +1 because sheet rows are 1-indexed and we skipped header
        data: {}
      };
      
      // Extract all column data
      Object.keys(columns).forEach(key => {
        if (columnIndices[key] !== undefined) {
          rowData.data[key] = row[columnIndices[key]];
        }
      });
      
      unprocessedRows.push(rowData);
    }
  }
  
  Logger.log(`Found ${unprocessedRows.length} unprocessed rows`);
  
  // Process each unprocessed row
  unprocessedRows.forEach(rowInfo => {
    try {
      const nameValue = rowInfo.data.name;
      
      // Check if name is valid according to validation rule
      if (!isValidName(nameValidation, nameValue, spreadsheet)) {
        // Set emailSent to "invalid name" for invalid names
        sheet.getRange(rowInfo.rowNumber, columnIndices.emailSent + 1).setValue("invalid name");
        Logger.log(`Invalid name for row ${rowInfo.rowNumber}, marked as "invalid name"`);
        return;
      }
      
      // Send email
      sendLateSheetEmail(rowInfo.data);
      
      // Update emailSent column with current timestamp
      const timestamp = new Date();
      sheet.getRange(rowInfo.rowNumber, columnIndices.emailSent + 1).setValue(timestamp);
      
      Logger.log(`Email sent and timestamp recorded for row ${rowInfo.rowNumber}`);
    } catch (error) {
      Logger.log(`Error processing row ${rowInfo.rowNumber}: ${error.toString()}`);
    }
  });
}

/**
 * Format date to show month and day
 * @param {Date|string|number} dateValue - The date value to format
 * @returns {string} Formatted date string
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (isNaN(date.getTime())) return dateValue.toString();
  return date.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
}

/**
 * Send email notification about late sheet entry
 * This function contains no Google Sheets logic
 * 
 * @param {Object} data - The row data containing late sheet information
 */
function sendLateSheetEmail(data) {
  const name = data.name || 'Member';
  const emailAddress = `${name}@directactioneverywhere.com`;
  const subject = "You've been added to the late sheet.";
  
  // Build email body with details
  let body = `Hello ${name},\n\n`;
  body += "You've been added to the late sheet with the following details:\n\n";
  
  if (data.date) {
    const formattedDate = formatDate(data.date);
    body += `Date: ${formattedDate}\n`;
  }
  if (data.violation) {
    body += `Violation: ${data.violation}\n`;
  }
  if (data.howLate) {
    body += `How Late: ${data.howLate}\n`;
  }
  if (data.notes) {
    body += `Notes: ${data.notes}\n`;
  }
  
  body += "\nBest regards,\nLate Sheet System";
  
  // Send the email
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: body
  });
}
