function divideEmissions() {
    // Open the active spreadsheet and select the "Carpool Data" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Carpool Data");
    var lastRow = sheet.getLastRow();
  
    // Loop through each row of the sheet
    for (var i = 2; i <= lastRow; i++) {  // Assuming the first row is headers
      Logger.log("Processing row: " + i);
      
      // Get the values from the relevant columns using carpoolMap constants
      var partialPassengerTokens = sheet.getRange(i, carpoolMap.partialPassengerTokens + 1).getValue(); // Column B
      var fullPassengerTokens = sheet.getRange(i, carpoolMap.fullPassengerTokens + 1).getValue();       // Column C
      var driverFirstCo2Value = sheet.getRange(i, carpoolMap.driverFirstCo2 + 1).getValue();            // Column D
      var driverSecondCo2Value = sheet.getRange(i, carpoolMap.driverSecondCo2 + 1).getValue();          // Column E
  
      // Convert token strings to arrays if they are stored as comma-separated values, handling empty strings
      var partialPassengerCount = partialPassengerTokens && partialPassengerTokens.trim() !== "" 
        ? partialPassengerTokens.split(',').length 
        : 0;
  
      var fullPassengerCount = fullPassengerTokens && fullPassengerTokens.trim() !== "" 
        ? fullPassengerTokens.split(',').length 
        : 0;
  
      // Log the retrieved values for debugging
      Logger.log("Partial Passenger Count (Row " + i + "): " + partialPassengerCount);
      Logger.log("Full Passenger Count (Row " + i + "): " + fullPassengerCount);
      Logger.log("Driver First CO2 Value (Row " + i + "): " + driverFirstCo2Value);
      Logger.log("Driver Second CO2 Value (Row " + i + "): " + driverSecondCo2Value);
      
      // Calculate sharedFirst by dividing driverFirstCo2Value among passengers and driver
      var sharedFirst = driverFirstCo2Value / (partialPassengerCount + 1); // +1 includes the driver
  
      Logger.log("Shared First Emissions (Row " + i + "): " + sharedFirst);
      sheet.getRange(i, carpoolMap.sharedFirst + 1).setValue(sharedFirst); // Set the result in the 'sharedFirst' column (Column F)
      
      // Calculate sharedSecond by dividing driverSecondCo2Value among passengers and driver
      var sharedSecond = driverSecondCo2Value / (fullPassengerCount + 1); // +1 includes the driver
  
      Logger.log("Shared Second Emissions (Row " + i + "): " + sharedSecond);
      sheet.getRange(i, carpoolMap.sharedSecond + 1).setValue(sharedSecond); // Set the result in the 'sharedSecond' column (Column G)
    }
    
    // Log completion
    Logger.log("Finished processing all rows.");
  }