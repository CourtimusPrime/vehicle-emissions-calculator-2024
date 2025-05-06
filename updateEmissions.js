function updateEmissionsCalculator() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const distanceCalcSheet = ss.getSheetByName('Distance Calculator');
    const emissionsCalcSheet = ss.getSheetByName('Emissions Calculator');
    const bniMembersSheet = ss.getSheetByName('BNI Members');
    const substitutesSheet = ss.getSheetByName('Substitutes');
    const visitorsSheet = ss.getSheetByName('Visitors');
    const carpoolSheet = ss.getSheetByName('Carpool Data');
  
    const distanceData = distanceCalcSheet.getDataRange().getValues();
    const emissionsData = emissionsCalcSheet.getDataRange().getValues();
    const bniMembersData = bniMembersSheet.getDataRange().getValues();
    const substitutesData = substitutesSheet.getDataRange().getValues();
    const visitorsData = visitorsSheet.getDataRange().getValues();
    const carpoolData = carpoolSheet.getDataRange().getValues();
  
    // Iterate over Distance Calculator rows (skip headers)
    for (let i = 1; i < distanceData.length; i++) {
      const row = distanceData[i];
      const token = row[distanceCalculator.token]; // Token from Distance Calculator
      const firstTravMode = row[distanceCalculator.startTravMode];
      const secondTravMode = row[distanceCalculator.finTravMode];
      const startDistance = row[distanceCalculator.firstDistance];
      const secondDistance = row[distanceCalculator.secondDistance];
  
      let member = findUserByToken(token, bniMembersData, substitutesData, visitorsData); // Find member data before using
  
      // Default values for firstFuel and secondFuel
      let firstFuel = '';
      let secondFuel = '';
  
      if (member) {
        // Retrieve carFuel and finalCarFuel from member data
        const carFuel = member.carFuel;
        const finalCarFuel = member.finalCarFuel;
  
        // Set firstFuel to carFuel if available
        if (carFuel) {
          firstFuel = carFuel;
          console.log(`Token ${token}: Copied carFuel "${carFuel}" to firstFuel.`);
        }
  
        // Set secondFuel to finalCarFuel if available
        if (finalCarFuel) {
          secondFuel = finalCarFuel;
          console.log(`Token ${token}: Copied finalCarFuel "${finalCarFuel}" to secondFuel.`);
        }
      }
  
      // If firstFuel is "Car", check if they are a passenger in a carpool
      if (firstFuel.toLowerCase() === 'car') {
        const driverFuel = findDriverFuelInCarpool(token, carpoolData, bniMembersData, substitutesData, visitorsData);
        if (driverFuel) {
          firstFuel = driverFuel;
          console.log(`Token ${token}: Passenger in carpool, copied driver's fuel type "${driverFuel}" to firstFuel.`);
        }
      }
  
      // If firstFuel is blank, copy secondFuel or fall back to firstTravMode
      if (!firstFuel) {
        firstFuel = secondFuel || firstTravMode;
        console.log(`Token ${token}: firstFuel was blank, copied "${secondFuel || firstTravMode}" instead.`);
      }
  
      // If secondFuel is still blank, fall back to secondTravMode
      if (!secondFuel) {
        secondFuel = secondTravMode;
        console.log(`Token ${token}: secondFuel was blank, copied "${secondTravMode}" instead.`);
      }
  
      // Check if the token exists in Emissions Calculator
      const emissionsRowIndex = findTokenInSheet(token, emissionsData);
  
      if (emissionsRowIndex !== -1) {
        // Token exists, check if distances match
        const emissionsRow = emissionsData[emissionsRowIndex];
        const existingStartDistance = emissionsRow[emissionsMap.startDistance];
        const existingSecondDistance = emissionsRow[emissionsMap.secondDistance];
  
        if (startDistance === existingStartDistance && secondDistance === existingSecondDistance) {
          console.log(`Token ${token}: Distances match, skipping row.`);
          continue;
        } else {
          // Distances don't match, update the distances
          emissionsCalcSheet.getRange(emissionsRowIndex + 1, emissionsMap.startDistance + 1).setValue(startDistance);
          emissionsCalcSheet.getRange(emissionsRowIndex + 1, emissionsMap.secondDistance + 1).setValue(secondDistance);
          console.log(`Token ${token}: Updated startDistance to "${startDistance}" and secondDistance to "${secondDistance}" in Emissions Calculator.`);
        }
      } else {
        // Token doesn't exist, proceed with creating a new row in Emissions Calculator
        // Initialize Emissions Calculator data array with potential nulls handled
        let emissionsRow = [token, startDistance, firstFuel, secondDistance, secondFuel, '', '', ''];
        emissionsCalcSheet.appendRow(emissionsRow);
        console.log(`Token ${token}: Added new row with startDistance "${startDistance}", firstFuel "${firstFuel}", secondDistance "${secondDistance}", and secondFuel "${secondFuel}" to Emissions Calculator.`);
      }
    }
  }
  
  // Helper function to find user's driver fuel in a carpool
  function findDriverFuelInCarpool(passengerToken, carpoolData, bniMembersData, substitutesData, visitorsData) {
    for (let i = 1; i < carpoolData.length; i++) {
      const row = carpoolData[i];
      const driverToken = row[carpoolMap.driverToken];
      const partialPassengers = row[carpoolMap.partialPassengerTokens] || [];
      const fullPassengers = row[carpoolMap.fullPassengerTokens] || [];
  
      // Check if the passenger token is in either partial or full passengers
      if (partialPassengers.includes(passengerToken) || fullPassengers.includes(passengerToken)) {
        // Find the driver's carFuel
        const driver = findUserByToken(driverToken, bniMembersData, substitutesData, visitorsData);
        if (driver) {
          console.log(`Found driver's token ${driverToken} for passenger ${passengerToken}, with fuel type "${driver.carFuel}".`);
          return driver.carFuel;
        }
      }
    }
    console.log(`No driver found for passenger token ${passengerToken} in Carpool Data.`);
    return null;
  }
  
  // Helper function to find user in Emissions Calculator by token
  function findTokenInSheet(token, emissionsData) {
    for (let i = 1; i < emissionsData.length; i++) {
      if (emissionsData[i][emissionsMap.token] === token) {
        return i;  // Return the row index where the token was found
      }
    }
    return -1;  // Return -1 if token was not found
  }
  
  // Helper function to find user in BNI Members, Substitutes, or Visitors by token
  function findUserByToken(token, bniMembersData, substitutesData, visitorsData) {
    const dataSources = [
      { data: bniMembersData, map: bniMap, source: 'BNI Members' },
      { data: substitutesData, map: subMap, source: 'Substitutes' },
      { data: visitorsData, map: visitorMap, source: 'Visitors' }
    ];
  
    for (let { data, map, source } of dataSources) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][map.token] === token) {
          const carFuel = data[i][map.carFuel];
          const finalCarFuel = data[i][map.finalCarFuel];
          console.log(`Token ${token}: Found in ${source} - carFuel: "${carFuel}", finalCarFuel: "${finalCarFuel}"`);
          return {
            carFuel: carFuel,
            finalCarFuel: finalCarFuel
          };
        }
      }
    }
    console.log(`Token ${token}: Not found in BNI Members, Substitutes, or Visitors.`);
    return null;
  }