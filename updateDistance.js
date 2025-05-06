function updateDistanceCalculator() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const frontSheet = ss.getSheetByName('Front Page');
    const bniSheet = ss.getSheetByName('BNI Members');
    const subSheet = ss.getSheetByName('Substitutes');
    const visitorSheet = ss.getSheetByName('Visitors');
    const distanceSheet = ss.getSheetByName('Distance Calculator');
    
    const bniData = bniSheet.getDataRange().getValues();
    const subData = subSheet.getDataRange().getValues();
    const visitorData = visitorSheet.getDataRange().getValues();
    const distanceData = distanceSheet.getDataRange().getValues();
  
    Logger.log(`BNI Data Rows: ${bniData.length}`);
    Logger.log(`Substitutes Data Rows: ${subData.length}`);
    Logger.log(`Visitors Data Rows: ${visitorData.length}`);
    Logger.log(`Distance Calculator Data Rows: ${distanceData.length}`);
    
    // Create a set of tokens that already exist in Distance Calculator
    let existingTokens = new Set();
    for (let i = 1; i < distanceData.length; i++) {
      Logger.log(`Existing token found in Distance Calculator: ${distanceData[i][distanceCalculator.token]}`);
      existingTokens.add(distanceData[i][distanceCalculator.token]);
    }
    Logger.log(`Existing tokens in Distance Calculator: ${[...existingTokens].join(', ')}`);
  
    // Function to transfer data to Distance Calculator
    function transferToDistanceCalculator(row, map, sheetName) {
      const token = row[map.token];
      Logger.log(`Processing row with token: ${token} from ${sheetName}`);
      
      // If token does not already exist in Distance Calculator
      if (!existingTokens.has(token)) {
        Logger.log(`Token ${token} not found in Distance Calculator, transferring data`);
        const newData = [
          token,                                   // Column A: Token
          row[map.startLocation] || '',            // Column B: Start Location
          row[map.transportMode] || '',            // Column C: Transport Mode
          row[map.firstStopLocation] || '',        // Column D: First Stop Location
          row[map.secondStopLocation] || '',       // Column E: Second Stop Location
          row[map.thirdStopLocation] || '',        // Column F: Third Stop Location
          row[map.fourthStopLocation] || '',       // Column G: Fourth Stop Location
          row[map.fifthStopLocation] || '',        // Column H: Fifth Stop Location
          row[map.sixthStopLocation] || '',        // Column I: Sixth Stop Location
          row[map.seventhStopLocation] || '',      // Column J: Seventh Stop Location
          row[map.eigthStopLocation] || '',        // Column K: Eighth Stop Location
          '', // Column L: distanceToFirstStop (initially blank)
          '', // Column M: distanceToSecondStop (initially blank)
          '', // Column N: distanceToThirdStop (initially blank)
          '', // Column O: distanceToFourthStop (initially blank)
          '', // Column P: distanceToConference (initially blank)
          '', // Column Q: conferenceLocation (initially blank)
          row[map.finalLocation] || '',            // Column R: Final Location (should be at index 17)
          row[map.finalTransMode] || row[map.transportMode] || '',  // Column S: Final Transport Mode (should be at index 18)
          '', // Column T: distanceToFinalLocation (initially blank)
          '', // Column U: organicDistance (initially blank)
          '', // Column V: firstDistance (initially blank)
          '', // Column W: secondDistance (initially blank)
          '', // Column X: totalDistance (initially blank)
        ];
  
        // Append the new data to Distance Calculator sheet
        distanceSheet.appendRow(newData);
        
        // Add token to the existingTokens set to avoid duplication
        existingTokens.add(token);
        
        Logger.log(`Successfully transferred token ${token} to Distance Calculator`);
      } else {
        Logger.log(`Token ${token} already exists in Distance Calculator, skipping transfer`);
      }
    }
  
    // Transfer data from BNI Members to Distance Calculator
    for (let i = 1; i < bniData.length; i++) {
      Logger.log(`Processing BNI Member row ${i}`);
      transferToDistanceCalculator(bniData[i], bniMap, 'BNI Members');
    }
  
    // Transfer data from Substitutes to Distance Calculator
    for (let i = 1; i < subData.length; i++) {
      Logger.log(`Processing Substitute row ${i}`);
      transferToDistanceCalculator(subData[i], subMap, 'Substitutes');
    }
  
    // Transfer data from Visitors to Distance Calculator
    for (let i = 1; i < visitorData.length; i++) {
      Logger.log(`Processing Visitor row ${i}`);
      transferToDistanceCalculator(visitorData[i], visitorMap, 'Visitors');
    }
  }