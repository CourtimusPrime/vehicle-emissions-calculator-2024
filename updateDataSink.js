function updateDataSink() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emissionsSheet = ss.getSheetByName('Emissions Calculator');
    const dataSinkSheet = ss.getSheetByName('dataSink');
    const distanceSheet = ss.getSheetByName('Distance Calculator');
    const substitutesSheet = ss.getSheetByName('Substitutes');
    const bniSheet = ss.getSheetByName('BNI Members');
    const visitorsSheet = ss.getSheetByName('Visitors');
  
    // Get all data from relevant sheets, skipping headers
    const emissionsData = emissionsSheet.getDataRange().getValues().slice(1);
    const dataSinkData = dataSinkSheet.getDataRange().getValues().slice(1);
    const distanceData = distanceSheet.getDataRange().getValues().slice(1);
    const substitutesData = substitutesSheet.getDataRange().getValues().slice(1);
    const bniData = bniSheet.getDataRange().getValues().slice(1);
    const visitorsData = visitorsSheet.getDataRange().getValues().slice(1);
  
    // Create a map of tokens to totalDistance from Distance Calculator
    const distanceMap = new Map();
    distanceData.forEach(row => {
      const token = row[distanceCalculator.token];
      const totalDistance = row[distanceCalculator.totalDistance];
      distanceMap.set(token, totalDistance);
    });
  
    // Create a set of tokens present in the subForToken column in the "Substitutes" sheet
    const subForTokensSet = new Set();
    substitutesData.forEach(row => {
      const subForToken = row[subMap.subForToken];
      if (subForToken) {
        subForTokensSet.add(subForToken);
      }
    });
  
    // Function to determine the type based on the token
    function getType(token) {
      if (bniData.some(row => row[bniMap.token] === token)) {
        return 'BNI Members';
      } else if (substitutesData.some(row => row[subMap.token] === token)) {
        return 'Substitutes';
      } else if (visitorsData.some(row => row[visitorMap.token] === token)) {
        return 'Visitors';
      }
      return 'Unknown';
    }
  
    // Prepare data to be written to the dataSink sheet
    const updatedDataSink = [];
  
    // Sync emissions data to dataSink
    emissionsData.forEach(row => {
      const token = row[emissionsMap.token];
      const firstEmission = row[emissionsMap.firstEmission];
      const secondEmission = row[emissionsMap.secondEmissions];
      const totalEmission = row[emissionsMap.totalEmissions];
  
      // Find the corresponding totalDistance using the token
      const totalDistance = distanceMap.get(token) || 0; // Default to 0 if not found
  
      // Check if the token is not in the subForTokensSet before adding it
      if (!subForTokensSet.has(token)) {
        // Determine the type of the entry
        const type = getType(token);
  
        // Prepare the row for the dataSink sheet
        const dataSinkRow = [
          token,
          type,
          firstEmission,
          secondEmission,
          totalEmission,
          totalDistance
        ];
  
        updatedDataSink.push(dataSinkRow);
      }
    });
  
    // Write the updated data to the dataSink sheet, starting from row 2 to skip the header
    if (updatedDataSink.length > 0) {
      dataSinkSheet.getRange(2, 1, updatedDataSink.length, updatedDataSink[0].length).setValues(updatedDataSink);
    }
  }