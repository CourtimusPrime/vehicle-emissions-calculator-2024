function transferData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const typeformSheet = ss.getSheetByName('Typeform Inputs');
    const bniSheet = ss.getSheetByName('BNI Members');
    const subSheet = ss.getSheetByName('Substitutes');
    const visitorSheet = ss.getSheetByName('Visitors');
  
    const inputData = typeformSheet.getDataRange().getValues();
    
    // Calculate the date range for the previous month
    const today = new Date();
    const startOfPreviousMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const endOfPreviousMonth = new Date(today.getFullYear(), today.getMonth(), 0);
  
    for (let i = 1; i < inputData.length; i++) { // Skip header row
      const row = inputData[i];
      const memberStatus = row[typeformMap.type];
      const token = row[typeformMap.token];
      const submitTime = new Date(row[typeformMap.submitTime]);
  
      // If monthlyLimit is true, skip rows outside of the previous month range.
      if (monthlyLimit === 'true' && (submitTime < startOfPreviousMonth || submitTime > endOfPreviousMonth)) {
        Logger.log(`Skipping row ${i + 1} - Submit time ${submitTime} is not within the last month`);
        continue;
      }
  
      Logger.log(`Processing row ${i + 1} with token: ${token}`);
  
      // Prepare data object for transfer
      let data = {
        token: token,
        emailAddress: row[typeformMap.emailAddress],
        joinDate: row[typeformMap.joinDate],
        startLocation: row[typeformMap.startLocation],
        firstStopLocation: row[typeformMap.stopLocation],
        secondStopLocation: row[typeformMap.secondStop],
        thirdStopLocation: row[typeformMap.thirdStop],
        fourthStopLocation: row[typeformMap.fourthStop],
        fifthStopLocation: row[typeformMap.fifthStop],
        sixthStopLocation: row[typeformMap.sixthStop],
        seventhStopLocation: row[typeformMap.seventhStop],
        eigthStopLocation: row[typeformMap.eigthStop],
        transportMode: row[typeformMap.transportMode],
        carBrand: row[typeformMap.carBrand],
        carModel: row[typeformMap.carModel],
        carFuel: row[typeformMap.carFuel],
        finalLocation: row[typeformMap.finalLocation] === 'Somewhere else' ? row[typeformMap.finalLocationElse] : row[typeformMap.startLocation],
        finalTransMode: row[typeformMap.finalTransportMode].trim().toLowerCase() === 'by careem' ? 'Car' : 
                        row[typeformMap.finalTransportMode].trim().toLowerCase() === 'by transit' ? 'Transit' : row[typeformMap.finalTransportMode],
        finalCarBrand: row[typeformMap.finalCarBrand],
        finalCarModel: row[typeformMap.finalCarModel],
        finalCarFuel: row[typeformMap.finalCarFuel]
      };
  
      // Find the carpool driver's token if available
      if (row[typeformMap.carpoolDriverFirstName] && row[typeformMap.carpoolDriverLastName]) {
        data.carpoolDriverUsername = findCarpoolDriverToken(inputData, row[typeformMap.carpoolDriverFirstName], row[typeformMap.carpoolDriverLastName]);
      } else {
        data.carpoolDriverUsername = '';
      }
  
      // New condition: If finalCar and finalTransMode fields are empty and transportMode is "Car", copy car details and transportMode
      if (!data.finalCarBrand && !data.finalCarModel && !data.finalCarFuel && !data.finalTransMode && data.transportMode.toLowerCase() === 'car') {
        Logger.log(`Copying car details and transport mode for token: ${data.token}`);
        data.finalCarBrand = data.carBrand;
        data.finalCarModel = data.carModel;
        data.finalCarFuel = data.carFuel;
        data.finalTransMode = data.transportMode;
      }
  
      // Additional condition: If finalTransMode is still empty, copy transportMode
      if (!data.finalTransMode) {
        Logger.log(`Final transport mode is empty for token: ${data.token}, copying transport mode: ${data.transportMode}`);
        data.finalTransMode = data.transportMode;
      }
  
      // Check the member status to determine the destination sheet
      if (memberStatus === 'BNI Member') {
        Logger.log(`Transferring to BNI Members: ${token}`);
        handleDataTransfer(data, bniSheet, bniMap);
      } else if (memberStatus === 'Substitute') {
        // Find the token of the person being subbed for
        const subbingForFirstName = row[typeformMap.subbingForFirstName].toLowerCase().trim();
        const subbingForLastName = row[typeformMap.subbingForLastName].toLowerCase().trim();
        const subForToken = findTokenByFullName(inputData, subbingForFirstName, subbingForLastName);
  
        if (subForToken) {
          data.subForToken = subForToken;
        } else {
          Logger.log(`No matching token found for ${subbingForFirstName} ${subbingForLastName}`);
        }
  
        data.subbingDays = row[typeformMap.subbingTotalDays];
        Logger.log(`Transferring to Substitutes: ${token}`);
        handleDataTransfer(data, subSheet, subMap);
      } else if (memberStatus === 'Visitor') {
        data.substituteDay = row[typeformMap.visitDate];
        Logger.log(`Transferring to Visitors: ${token}`);
        handleDataTransfer(data, visitorSheet, visitorMap);
      } else {
        Logger.log(`Unrecognized member status for row ${i + 1}: ${memberStatus}`);
      }
    }
  }
  
  // Find a token for the carpool driver based on first name and last name in the Typeform Inputs data
  function findCarpoolDriverToken(inputData, firstName, lastName) {
    for (let i = 1; i < inputData.length; i++) { // Skip header row
      const row = inputData[i];
      const rowFirstName = row[typeformMap.firstName].toLowerCase().trim();
      const rowLastName = row[typeformMap.lastName].toLowerCase().trim();
      const rowToken = row[typeformMap.token];
  
      if (rowFirstName === firstName.toLowerCase().trim() && rowLastName === lastName.toLowerCase().trim()) {
        return rowToken;
      }
    }
    return '';
  }
  
  // Find a token based on first name and last name in the Typeform Inputs data
  function findTokenByFullName(inputData, firstName, lastName) {
    for (let i = 1; i < inputData.length; i++) { // Skip header row
      const row = inputData[i];
      const rowFirstName = row[typeformMap.firstName].toLowerCase().trim();
      const rowLastName = row[typeformMap.lastName].toLowerCase().trim();
  
      if (rowFirstName === firstName && rowLastName === lastName) {
        return row[typeformMap.token];
      }
    }
    return null;
  }
  
  // Handle data transfer to the target sheet
  function handleDataTransfer(data, targetSheet, targetMap) {
    const targetData = targetSheet.getDataRange().getValues();
    let found = false;
  
    // Check if the token already exists and update it
    for (let j = 1; j < targetData.length; j++) {
      if (targetData[j][targetMap.token] === data.token) {
        Logger.log(`Token ${data.token} already exists in ${targetSheet.getName()}, skipping transfer`);
        found = true;
        break;
      }
    }
  
    // If token not found, append new data
    if (!found) {
      const lastRow = targetSheet.getLastRow() + 1;
      Logger.log(`Appending new record for token: ${data.token} in ${targetSheet.getName()}`);
      transferRowData(lastRow, data, targetSheet, targetMap);
    }
  }
  
  // Transfer data to the appropriate row in the target sheet
  function transferRowData(targetRowIndex, data, targetSheet, targetMap) {
    let values = [];
  
    // Adjust the values array based on the target sheet name
    if (targetSheet.getName() === 'BNI Members') {
      values = [
        data.token,
        data.emailAddress,
        data.joinDate,
        data.startLocation,
        data.firstStopLocation,
        data.secondStopLocation || '',
        data.thirdStopLocation || '',
        data.fourthStopLocation || '',
        data.transportMode,
        data.carpoolDriverUsername || '',
        data.carBrand || '',
        data.carModel || '',
        data.carFuel || '',
        data.finalLocation,
        data.finalTransMode,
        data.finalCarBrand || '',
        data.finalCarModel || '',
        data.finalCarFuel || '',
        data.fifthStopLocation || '',
        data.sixthStopLocation || '',
        data.seventhStopLocation || '',
        data.eigthStopLocation || ''
      ];
    } else if (targetSheet.getName() === 'Substitutes') {
      values = [
        data.token,
        data.emailAddress,
        data.subForToken || '',
        data.subbingDays || '',
        data.startLocation,
        data.firstStopLocation,
        data.secondStopLocation || '',
        data.thirdStopLocation || '',
        data.fourthStopLocation || '',
        data.transportMode,
        data.carpoolDriverUsername || '',
        data.carBrand || '',
        data.carModel || '',
        data.carFuel || '',
        data.finalLocation,
        data.finalTransMode,
        data.finalCarBrand || '',
        data.finalCarModel || '',
        data.finalCarFuel || '',
        data.fifthStopLocation || '',
        data.sixthStopLocation || '',
        data.seventhStopLocation || '',
        data.eigthStopLocation || ''
      ];
    } else if (targetSheet.getName() === 'Visitors') {
      values = [
        data.token,
        data.emailAddress,
        data.substituteDay || '',
        data.startLocation,
        data.firstStopLocation,
        data.secondStopLocation || '',
        data.thirdStopLocation || '',
        data.fourthStopLocation || '',
        data.transportMode,
        data.carpoolDriverUsername || '',
        data.carBrand || '',
        data.carModel || '',
        data.carFuel || '',
        data.finalLocation,
        data.finalTransMode,
        data.finalCarBrand || '',
        data.finalCarModel || '',
        data.finalCarFuel || '',
        data.fifthStopLocation || '',
        data.sixthStopLocation || '',
        data.seventhStopLocation || '',
        data.eigthStopLocation || ''
      ];
    }
  
    Logger.log(`Writing data to row ${targetRowIndex} in ${targetSheet.getName()}: ${values.join(", ")}`);
    targetSheet.getRange(targetRowIndex, 1, 1, values.length).setValues([values]);
  }