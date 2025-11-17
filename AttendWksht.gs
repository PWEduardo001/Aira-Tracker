/**
 * @OnlyCurrentDoc
 *
 * Counts unique attendees for each date specified in the 'Attend Wksht' tab.
 * This version is updated to handle dates that are formatted as PLAIN TEXT in the source sheets.
 * It sources its data from the 'online attendance' tab, creating a
 * unique attendee by combining their first and last names.
 */
function countUniqueOnlineAttendees() {
  Logger.log('Script execution started.');

  // --- Configuration ---
  const summarySheetName = 'Attend Wksht';
  const attendanceSheetName = 'online attendance';

  // --- Get the active spreadsheet and required sheets ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(summarySheetName);
  const attendanceSheet = ss.getSheetByName(attendanceSheetName);

  // --- Validate that the sheets exist ---
  if (!summarySheet) {
    Logger.log(`Error: The summary sheet named "${summarySheetName}" was not found.`);
    return;
  }
  if (!attendanceSheet) {
    Logger.log(`Error: The attendance data sheet named "${attendanceSheetName}" was not found.`);
    return;
  }

  // --- 1. Process the 'online attendance' data ---
  Logger.log(`Reading data from "${attendanceSheetName}" sheet.`);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  const attendeesByDate = new Map();

  // Define column indexes (0-based) for the 'online attendance' sheet
  const firstNameCol = 2; // Column C
  const lastNameCol = 3;  // Column D
  const dateCol = 4;      // Column E

  // Loop through attendance data, skipping the header row (i = 1)
  for (let i = 1; i < attendanceData.length; i++) {
    const row = attendanceData[i];
    const firstName = row[firstNameCol];
    const lastName = row[lastNameCol];
    const dateText = row[dateCol]; // Read the date as whatever it is (text or date)

    // Try to convert the text into a valid date object.
    const dateObj = new Date(dateText);

    // Check if we have names AND the date text was successfully converted into a valid date.
    // !isNaN(dateObj.getTime()) is a reliable way to check for a valid date.
    if (firstName && lastName && dateText && !isNaN(dateObj.getTime())) {
      const fullName = `${firstName.toString().trim()} ${lastName.toString().trim()}`.toLowerCase();
      
      // Format the valid date object into a consistent string 'yyyy-MM-dd' to use as a key
      const dateString = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

      if (!attendeesByDate.has(dateString)) {
        attendeesByDate.set(dateString, new Set());
      }
      
      attendeesByDate.get(dateString).add(fullName);
    }
  }
  Logger.log('Finished processing raw attendance data.');

  // --- 2. Get dates from 'Attend Wksht' and calculate counts ---
  Logger.log(`Processing dates from "${summarySheetName}" sheet.`);
  // Start from row 2, column 1 (A2), and go to the last row
  const lastRow = summarySheet.getLastRow();
  if (lastRow < 2) {
      Logger.log('No dates found in the summary sheet to process.');
      Logger.log('Script execution completed.');
      return;
  }
  const summaryDataRange = summarySheet.getRange('A2:A' + lastRow);
  const summaryDates = summaryDataRange.getValues();

  const countsToWrite = [];

  // Loop through each date in the summary sheet
  for (let i = 0; i < summaryDates.length; i++) {
    const dateText = summaryDates[i][0];
    let uniqueCount = 0;

    const dateObj = new Date(dateText);

    // Check if the date text from the summary sheet is also a valid date
    if (dateText && !isNaN(dateObj.getTime())) {
      const dateString = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      
      if (attendeesByDate.has(dateString)) {
        uniqueCount = attendeesByDate.get(dateString).size;
      }
    }
    
    countsToWrite.push([uniqueCount]);
  }

  // --- 3. Write the counts back to the 'Attend Wksht' sheet ---
  if (countsToWrite.length > 0) {
    Logger.log(`Writing ${countsToWrite.length} counts to Column B.`);
    const targetRange = summarySheet.getRange(2, 2, countsToWrite.length, 1); // Range starts at B2
    targetRange.setValues(countsToWrite);
    Logger.log('Successfully updated attendance counts.');
  } else {
    Logger.log('No valid dates found in the summary sheet to process.');
  }

  Logger.log('Script execution completed.');
}
