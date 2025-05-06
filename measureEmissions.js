function calculateEmissions() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emissions Calculator');
    const data = sheet.getDataRange().getValues();
  
    const emissionFactors = {
      'petrol': 0.17,
      'diesel': 0.332,
      'gasoline': 0.306,
      'electric': 0.101,
      'hybrid': 0.1137,
      'bus': 1.016,
      'transit': 0.0271,
    };
  
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
      const row = data[i];
  
      // Extract necessary data from the row
      const startDistance = parseFloat(row[emissionsMap.startDistance]) || 0;
      const firstCarFuel = row[emissionsMap.firstCarFuel] ? row[emissionsMap.firstCarFuel].toLowerCase() : '';
      const firstMode = row[emissionsMap.firstMode] ? row[emissionsMap.firstMode].toLowerCase() : '';
  
      let firstEmissions = 0;
  
      // Calculate firstEmissions
      if (startDistance > 0) {
        if (emissionFactors[firstCarFuel]) {
          firstEmissions = startDistance * emissionFactors[firstCarFuel];
        } else if (emissionFactors[firstMode]) {
          firstEmissions = startDistance * emissionFactors[firstMode];
        } else {
          // Default emission factor if no match is found
          firstEmissions = 0;
        }
      }
  
      // Write firstEmissions back to the sheet
      sheet.getRange(i + 1, emissionsMap.firstEmission + 1).setValue(firstEmissions);
  
      // Extract data for second trip
      const secondDistance = parseFloat(row[emissionsMap.secondDistance]) || 0;
      const secondCarFuel = row[emissionsMap.secondCarFuel] ? row[emissionsMap.secondCarFuel].toLowerCase() : '';
      const secondMode = row[emissionsMap.secondMode] ? row[emissionsMap.secondMode].toLowerCase() : '';
  
      let secondEmissions = 0;
  
      // Calculate secondEmissions with respect to the sign of secondDistance
      if (secondDistance !== 0) {
        if (emissionFactors[secondCarFuel]) {
          secondEmissions = secondDistance * emissionFactors[secondCarFuel];
        } else if (emissionFactors[secondMode]) {
          secondEmissions = secondDistance * emissionFactors[secondMode];
        } else {
          // Default emission factor if no match is found
          secondEmissions = 0;
        }
      }
  
      // Write secondEmissions back to the sheet
      sheet.getRange(i + 1, emissionsMap.secondEmissions + 1).setValue(secondEmissions);
  
      // Calculate totalEmissions
      const totalEmissions = firstEmissions + secondEmissions;
  
      // Write totalEmissions to the sheet
      sheet.getRange(i + 1, emissionsMap.totalEmissions + 1).setValue(totalEmissions);
    }
  }