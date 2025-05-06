let monthlyLimit = 'true';

function automationTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const frontPageSheet = ss.getSheetByName('Front Page');
  automationBox = frontPageSheet.getRange('D12');
 // absenseBox = frontPageSheet.getRange('E20');
  resetBox = frontPageSheet.getRange('H12');
  automationStatus = frontPageSheet.getRange('F15');
  monthlyLimitBox = frontPageSheet.getRange('D14');
  if (monthlyLimitBox === true) {
    monthlyLimit = 'true';
  } else {
    monthlyLimit = 'false';
  }

  Logger.log('Monthly limit is set to:  ${monthlyLimit}');

  if (automationBox.getValue() === true) {
    automationBox.setValue(false);
    manualAutomation();
  //} //else if (absenseBox.getValue() === true) {
   // absenseBox.setValue(false);
   // triggerAbsense();
  } else if (resetBox.getValue() === true) {
    resetBox.setValue(false);
    resetSheets();
  }
}

function triggerAbsense() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const frontPageSheet = ss.getSheetByName('Front Page');

  // Extract the necessary details from the Front Page
  const firstName = frontPageSheet.getRange('C18').getValue();
  const lastName = frontPageSheet.getRange('E18').getValue();
  const date = frontPageSheet.getRange('G18').getValue();

  // Clear the cells C18, E18, and G18 after processing
  frontPageSheet.getRange('C18').setValue('');
  frontPageSheet.getRange('E18').setValue('');
  frontPageSheet.getRange('G18').setValue('');

  // Process the absentee information using conjoined lowercase username
  processAbsentee(firstName, lastName, date);
}

function processAbsentee(firstName, lastName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('Attendance Sheet');

  // Create username from first and last name in lowercase
  const username = (firstName + lastName).toLowerCase();

  // Find the username row in the Attendance Sheet
  const attendanceData = attendanceSheet.getDataRange().getValues();
  for (let j = 1; j < attendanceData.length; j++) { // Skip header row
    if (attendanceData[j][0].toLowerCase() === username) {
      Logger.log(`Found username: ${username} in row ${j + 1}`);
      
      // Get the current value in column D (index 3 for 4th column)
      let currentAbsentCount = parseInt(attendanceData[j][3], 10) || 0; // Default to 0 if empty or invalid

      // Increment the value in column D by 1
      currentAbsentCount += 1;
      Logger.log(`Updating absent count for ${username} to ${currentAbsentCount}`);

      // Update the value in column D
      attendanceSheet.getRange(j + 1, 4).setValue(currentAbsentCount); // Column D is the 4th column

      // Save changes back to the spreadsheet
      SpreadsheetApp.flush();
      Logger.log(`Absent count updated. Now calculating attendance totals and updating the dashboard.`);

      // Call the other functions
      calculateAttendanceTotals();
      updateDashboard();
      dashboardFunction();

      return;
    }
  }
  Logger.log('Username not found for the given name combination.');
}

function resetSheets () {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
    
  // List of sheets to clear
  const sheetsToClear = [
    'Attendance Sheet', 
    'Dashboard', 
    'Emissions Calculator', 
    'Carpool Data', 
    'BNI Members', 
    'Substitutes', 
    'Distance Calculator',
    'Visitors',
    'Monthly Capture'
  ];

  // Iterate through each sheet and clear all rows except the header (row 1)
  sheetsToClear.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {  // Ensure there are rows to clear besides the header
        sheet.deleteRows(2, lastRow - 1);  // Clear all rows except the header
      }
    }
  });
}

function manualAutomation() {
  // ============ UPDATE ARCHIVE ============
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('C16:I16').setBackground('white');
  automationStatus.setValue("Updating archive...");
  updateArchive()
  automationStatus.setValue("Archive updated.");
  // ============ CLEAR SHEETS =============
  automationStatus.setValue("Clearing processed data...");
  resetSheets();
  automationStatus.setValue("Cleared the sheets.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('C16').setBackground('green');
  // ============ TRANSFER DATA =============
  automationStatus.setValue("Distributing data...");
  transferData();
  automationStatus.setValue("Successfully distributed data.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('D16').setBackground('green');
  // ============ UPDATE DISTANCE CALCULATOR 
  automationStatus.setValue("Transfering data to Distance Calculator...");
  updateDistanceCalculator();
  automationStatus.setValue("Successfully transfered data.");
  automationStatus.setValue('Starting up distance calculator...');
  // ============ CALCULATE DISTANCES ========
  automationStatus.setValue("Measuring distances...");
  distanceFunction();
  automationStatus.setValue("Successfully measured distances.");
  automationStatus.setValue('Measuring distances...');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('E16').setBackground('green');
  // ============ UPDATE EMISSIONS CALCULATOR
  automationStatus.setValue("Updating Emissions Calculator...");
  updateEmissionsCalculator();
  automationStatus.setValue("Successfully updated Emissions Calculator.");
  automationStatus.setValue('Starting up emission calculator...');
  // ============ CALCULATE EMISSIONS =======
  automationStatus.setValue("Calculating emissions...");
  calculateEmissions();
  automationStatus.setValue("Successfully calculated emissions.");
  automationStatus.setValue('Measuring emissions...');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('F16').setBackground('green');
  // ============ UPDATE CARPOOL DATA =======
  automationStatus.setValue("Updating Carpool Data...");
  updateCarpoolData();
  automationStatus.setValue("Successfully updated Carpool Data.");
  automationStatus.setValue('Starting up carpool function...');
  // ============ ADJUST CARPOOL DATA =======
  automationStatus.setValue("Calculating carpool data...");
  divideEmissions();
  automationStatus.setValue("Successfully calculated carpool data.");
  automationStatus.setValue('Processing carpool data...');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('G16').setBackground('green');
  // ============ UPDATE MONTHLY CAPTURE ====
  automationStatus.setValue("Updating data sink...");
  updateDataSink();
  automationStatus.setValue("Successfully updated data sink");
  automationStatus.setValue('Snycing data sink...');
  handleDataSink();
  automationStatus.setValue('Succesfully synced data sink with Carpool Data.');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('H16').setBackground('green');
  // ============ UPDATE RECORD =============
  automationStatus.setValue("Updating monthly record...");
  updateMonthlyRecord();
  automationStatus.setValue("Successfully updated monthly records.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Front Page').getRange('I16').setBackground('green');
  automationStatus.setValue('Complete!');
}