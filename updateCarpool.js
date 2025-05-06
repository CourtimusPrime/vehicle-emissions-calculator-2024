function updateCarpoolData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const typeformSheet = ss.getSheetByName('Typeform Inputs');
    const bniSheet = ss.getSheetByName('BNI Members');
    const subSheet = ss.getSheetByName('Substitutes');
    const visitorSheet = ss.getSheetByName('Visitors');
    const carpoolDataSheet = ss.getSheetByName('Carpool Data');
    const emissionsSheet = ss.getSheetByName('Emissions Calculator');
  
    const typeformData = typeformSheet.getDataRange().getValues();
    const bniData = bniSheet.getDataRange().getValues();
    const subData = subSheet.getDataRange().getValues();
    const visitorData = visitorSheet.getDataRange().getValues();
    const carpoolData = carpoolDataSheet.getDataRange().getValues();
    const emissionsData = emissionsSheet.getDataRange().getValues();
  
    Logger.log('Loaded data from sheets. Starting processing...');
  
    // Iterate through each driver's data in Typeform Inputs
    typeformData.slice(1).forEach((row) => {
      const driverToken = row[typeformMap.token];
      const isCarpoolDriver = row[typeformMap.carpoolDriverStatus] === true || row[typeformMap.carpoolDriverStatus] === 'TRUE';
  
      Logger.log(`Processing driver token: ${driverToken}, isCarpoolDriver: ${isCarpoolDriver}`);
  
      // Only proceed if the entry is marked as a carpool driver
      if (driverToken && isCarpoolDriver) {
        // Retrieve emissions for the driver
        const emissionsRowIndex = findToken(emissionsData, driverToken, emissionsMap.token);
        let driverFirstCo2 = 0;
        let driverSecondCo2 = 0;
        if (emissionsRowIndex !== -1) {
          const rawFirstEmission = emissionsData[emissionsRowIndex - 1][emissionsMap.firstEmission];
          const rawSecondEmission = emissionsData[emissionsRowIndex - 1][emissionsMap.secondEmissions];
          Logger.log(`Raw firstEmission value: ${rawFirstEmission}`);
          Logger.log(`Raw secondEmission value: ${rawSecondEmission}`);
          driverFirstCo2 = parseFloat(rawFirstEmission) || 0;
          driverSecondCo2 = parseFloat(rawSecondEmission) || 0;
          Logger.log(`Parsed driverFirstCo2 value: ${driverFirstCo2}`);
          Logger.log(`Parsed driverSecondCo2 value: ${driverSecondCo2}`);
        }
  
        let carpoolIndex = findToken(carpoolData, driverToken, carpoolMap.driverToken);
        if (carpoolIndex === -1) {
          Logger.log(`No existing carpool entry for driver token: ${driverToken}, creating a new entry.`);
          const lastRow = carpoolDataSheet.getLastRow() + 1;
          carpoolDataSheet.getRange(lastRow, carpoolMap.driverToken + 1).setValue(driverToken);
          carpoolDataSheet.getRange(lastRow, carpoolMap.driverFirstCo2 + 1).setValue(driverFirstCo2);
          carpoolDataSheet.getRange(lastRow, carpoolMap.driverSecondCo2 + 1).setValue(driverSecondCo2);
          carpoolIndex = lastRow;
        } else {
          // Update existing row with CO2 values
          carpoolDataSheet.getRange(carpoolIndex, carpoolMap.driverFirstCo2 + 1).setValue(driverFirstCo2);
          carpoolDataSheet.getRange(carpoolIndex, carpoolMap.driverSecondCo2 + 1).setValue(driverSecondCo2);
        }
  
        let fullPassengerTokens = [];
        let partialPassengerTokens = [];
  
        // Search for matches in BNI Members, Substitutes, and Visitors
        [bniData, subData, visitorData].forEach((sheetData, sheetIndex) => {
          Logger.log(`Searching in ${['BNI Members', 'Substitutes', 'Visitors'][sheetIndex]}`);
          sheetData.slice(1).forEach((memberRow) => {
            const memberDriverToken = memberRow[bniMap.carpoolDriverUsername] || memberRow[subMap.carpoolDriverUsername] || memberRow[visitorMap.carpoolDriverUsername];
            const memberToken = memberRow[bniMap.token] || memberRow[subMap.token] || memberRow[visitorMap.token];
            const carpoolAfterStatus = memberRow[bniMap.carpoolAfterStatus] || memberRow[subMap.carpoolAfterStatus] || memberRow[visitorMap.carpoolAfterStatus];
  
            if (memberDriverToken === driverToken) {
              Logger.log(`Found matching carpoolDriverToken: ${driverToken} with token: ${memberToken}`);
              if (carpoolAfterStatus === true || carpoolAfterStatus === 'TRUE') {
                fullPassengerTokens.push(memberToken);
              } else {
                partialPassengerTokens.push(memberToken);
              }
            }
          });
        });
  
        updatePassengerTokens(carpoolDataSheet, carpoolIndex, partialPassengerTokens, fullPassengerTokens);
      }
    });
  
    Logger.log('Finished processing all entries.');
  }
  
  function updatePassengerTokens(sheet, rowNumber, partialTokens, fullTokens) {
    // Update partial and full passenger tokens in the specified row as JSON strings
    sheet.getRange(rowNumber, carpoolMap.partialPassengerTokens + 1).setValue(JSON.stringify(partialTokens));
    sheet.getRange(rowNumber, carpoolMap.fullPassengerTokens + 1).setValue(JSON.stringify(fullTokens));
    Logger.log(`Updated passenger tokens at row ${rowNumber}: Partial = ${JSON.stringify(partialTokens)}, Full = ${JSON.stringify(fullTokens)}`);
  }
  
  function findToken(data, token, tokenColumn) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][tokenColumn] === token) {
        Logger.log(`Found token ${token} at row ${i + 1}`);
        return i + 1;
      }
    }
    Logger.log(`Token ${token} not found`);
    return -1;
  }