function updateArchive() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSinkSheet = ss.getSheetByName('dataSink');
    const archiveSheet = ss.getSheetByName('Archive');
    const typeformSheet = ss.getSheetByName('Typeform Inputs');
  
    // Check if sheets exist
    if (!dataSinkSheet) {
      Logger.log('dataSink sheet not found.');
      return;
    }
    if (!archiveSheet) {
      Logger.log('Archive sheet not found.');
      return;
    }
    if (!typeformSheet) {
      Logger.log('Typeform Inputs sheet not found.');
      return;
    }
  
    // Get all data from the sheets, skipping the header row
    const dataSinkData = dataSinkSheet.getDataRange().getValues().slice(1); // Skip header row
    const typeformData = typeformSheet.getDataRange().getValues().slice(1); // Skip header row
  
    // Prepare a map of tokens to names and submission dates from Typeform Inputs for easy lookup
    const nameMap = new Map();
    const firstNameCol = typeformMap.firstName;
    const lastNameCol = typeformMap.lastName;
    const tokenCol = typeformMap.token;
    const submitTimeCol = typeformMap.submitTime;
  
    typeformData.forEach(row => {
      const token = row[tokenCol];
      const firstName = row[firstNameCol];
      const lastName = row[lastNameCol];
      const fullName = `${firstName} ${lastName}`;
      const submitTime = row[submitTimeCol] || "Unknown";
      nameMap.set(token, { fullName, submitTime });
    });
  
    // Prepare the data to be added to the Archive
    const archiveData = [];
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'); // Format date as 'yyyy-MM-dd'
    const timeString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'HH:mm:ss'); // Format time as 'HH:mm:ss'
  
    dataSinkData.forEach(row => {
      const token = row[dataSinkMap.token];
      const type = row[dataSinkMap.type];
      const totalEmission = row[dataSinkMap.totalEmission];
      const totalDistance = row[dataSinkMap.totalDistance];
  
      // Get the fullName and submitTime using the token, default to "Unknown" if not found
      const { fullName = "Unknown", submitTime = "Unknown" } = nameMap.get(token) || {};
  
      // Prepare the row for the Archive sheet with the current date and time
      const archiveRow = [
        dateString,    // Column A - dateHandled
        timeString,    // Column B - timeHandled
        fullName,      // Column C - fullName
        token,         // Column D - token
        submitTime,    // Column E - dateSubmitted
        type,          // Column F - type
        totalEmission, // Column G - totalEmission
        totalDistance  // Column H - totalDistance
      ];
  
      archiveData.push(archiveRow);
    });
  
    // Write the prepared data to the Archive sheet, starting from the first available row below existing data
    if (archiveData.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, archiveData.length, archiveData[0].length).setValues(archiveData);
      Logger.log('Data has been successfully archived.');
    } else {
      Logger.log('No data found to archive.');
    }
  }