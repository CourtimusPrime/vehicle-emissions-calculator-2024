function handleDataSink() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSinkSheet = ss.getSheetByName('dataSink');
    const carpoolDataSheet = ss.getSheetByName('Carpool Data');
  
    // Get all data from relevant sheets, skipping headers
    const dataSinkData = dataSinkSheet.getDataRange().getValues().slice(1);
    const carpoolData = carpoolDataSheet.getDataRange().getValues().slice(1);
  
    Logger.log('Data read from Carpool Data and dataSink sheets.');
  
    // Create a map of dataSink for quick lookups
    const dataSinkMapIndex = new Map(dataSinkData.map((row, index) => [row[dataSinkMap.token], index]));
  
    // Iterate over each row in the carpoolData to update dataSink
    carpoolData.forEach((carpoolRow, rowIndex) => {
      const driverToken = carpoolRow[carpoolMap.driverToken];
      const sharedFirst = carpoolRow[carpoolMap.sharedFirst] || 0;
      const sharedSecond = carpoolRow[carpoolMap.sharedSecond] || 0;
  
      Logger.log(`Processing Carpool Data Row ${rowIndex + 2}: driverToken = ${driverToken}, sharedFirst = ${sharedFirst}, sharedSecond = ${sharedSecond}`);
  
      // 1. Update the driver's emissions in dataSink
      if (dataSinkMapIndex.has(driverToken)) {
        const dataSinkRowIndex = dataSinkMapIndex.get(driverToken);
        const dataSinkRow = dataSinkData[dataSinkRowIndex];
  
        Logger.log(`Found driver token ${driverToken} in dataSink at row ${dataSinkRowIndex + 2}.`);
        
        // Update the driver's emissions based on sharedFirst and sharedSecond
        dataSinkRow[dataSinkMap.firstEmission] = sharedFirst;
        dataSinkRow[dataSinkMap.secondEmission] = sharedSecond;
  
        Logger.log(`Updated in-memory: Setting firstEmission (${sharedFirst}) and secondEmission (${sharedSecond}) for driverToken: ${driverToken} (Row: ${dataSinkRowIndex + 2})`);
      } else {
        Logger.log(`Driver token ${driverToken} not found in dataSink.`);
      }
  
      // Continue with passenger updates (these were previously working fine)
      const partialPassengers = Array.isArray(carpoolRow[carpoolMap.partialPassengerTokens]) 
        ? carpoolRow[carpoolMap.partialPassengerTokens] 
        : JSON.parse(carpoolRow[carpoolMap.partialPassengerTokens] || '[]');
      const fullPassengers = Array.isArray(carpoolRow[carpoolMap.fullPassengerTokens]) 
        ? carpoolRow[carpoolMap.fullPassengerTokens] 
        : JSON.parse(carpoolRow[carpoolMap.fullPassengerTokens] || '[]');
  
      partialPassengers.forEach(token => {
        if (dataSinkMapIndex.has(token)) {
          const dataSinkRowIndex = dataSinkMapIndex.get(token);
          const dataSinkRow = dataSinkData[dataSinkRowIndex];
          dataSinkRow[dataSinkMap.firstEmission] = sharedFirst;
          Logger.log(`Updated in-memory: Setting firstEmission (${sharedFirst}) for partial passenger token: ${token} (Row: ${dataSinkRowIndex + 2})`);
        }
      });
  
      fullPassengers.forEach(token => {
        if (dataSinkMapIndex.has(token)) {
          const dataSinkRowIndex = dataSinkMapIndex.get(token);
          const dataSinkRow = dataSinkData[dataSinkRowIndex];
          dataSinkRow[dataSinkMap.firstEmission] = sharedFirst;
          dataSinkRow[dataSinkMap.secondEmission] = sharedSecond;
          Logger.log(`Updated in-memory: Setting firstEmission (${sharedFirst}) and secondEmission (${sharedSecond}) for full passenger token: ${token} (Row: ${dataSinkRowIndex + 2})`);
        }
      });
    });
  
    // 4. Calculate totalEmission as the sum of firstEmission and secondEmission for each row
    dataSinkData.forEach((row, index) => {
      const firstEmission = row[dataSinkMap.firstEmission] || 0;
      const secondEmission = row[dataSinkMap.secondEmission] || 0;
      row[dataSinkMap.totalEmission] = firstEmission + secondEmission;
      Logger.log(`Calculating totalEmission for token ${row[dataSinkMap.token]}: firstEmission = ${firstEmission}, secondEmission = ${secondEmission}, totalEmission = ${row[dataSinkMap.totalEmission]} (Row: ${index + 2})`);
    });
  
    // Verify data before writing back
    Logger.log(`DataSink data before writing back: ${JSON.stringify(dataSinkData.slice(0, 5))}...`); // Show first 5 rows for brevity
  
    // Write back the updated data to the dataSink sheet, starting from row 2 to skip the header
    if (dataSinkData.length > 0) {
      dataSinkSheet.getRange(2, 1, dataSinkData.length, dataSinkData[0].length).setValues(dataSinkData);
      Logger.log(`Data successfully written back to dataSink sheet: Row count = ${dataSinkData.length}, Column count = ${dataSinkData[0].length}`);
    } else {
      Logger.log('No data to write back to dataSink sheet.');
    }
  }