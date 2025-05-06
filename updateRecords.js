function updateMonthlyRecord() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSinkSheet = ss.getSheetByName('dataSink');
    const monthlyRecordSheet = ss.getSheetByName('Monthly Record');
  
    // Check if sheets exist
    if (!dataSinkSheet) {
      Logger.log('dataSink sheet not found.');
      return;
    }
    if (!monthlyRecordSheet) {
      Logger.log('Monthly Record sheet not found.');
      return;
    }
  
    const dataSinkData = dataSinkSheet.getDataRange().getValues().slice(1); // Skip header row
    const monthlyRecordData = monthlyRecordSheet.getDataRange().getValues();
  
    const monthlyLimit = true; // Set this to 'true' for previous month only, 'false' for all data.
  
    // Get the previous month (not current) for the record
    const today = new Date();
    const previousMonthDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const previousMonth = previousMonthDate.toLocaleDateString('en-US', { month: 'long' });
  
    // Adjust month label based on monthlyLimit
    const monthLabel = monthlyLimit ? previousMonth : `All - ${previousMonth}`;
  
    // Check if the previous month already exists in the Monthly Record sheet
    const monthExists = monthlyRecordData.some(row => row[0] === monthLabel);
    if (monthExists) {
      Logger.log(`Record for ${monthLabel} already exists. Skipping update.`);
      return; // Exit the function if the month already exists
    }
  
    let totalPpl = 0, sumMem = 0, sumSub = 0, sumVis = 0;
    let distances = [], co2Emissions = [];
  
    // Iterate over dataSink data (skip header)
    for (let i = 0; i < dataSinkData.length; i++) {
      const row = dataSinkData[i];
      const type = row[dataSinkMap.type];           // Type (Member, Substitute, Visitor)
      const totalDistance = parseFloat(row[dataSinkMap.totalDistance]) || 0;  // Total Distance
      const totalEmission = parseFloat(row[dataSinkMap.totalEmission]) || 0;  // Total Emissions
  
      // Increment total people count
      totalPpl++;
  
      // Tally types based on rows
      if (type === 'BNI Members') {
        sumMem++;
      } else if (type === 'Substitutes') {
        sumSub++;
      } else if (type === 'Visitors') {
        sumVis++;
      }
  
      // Add distances and CO2 emissions to their respective arrays for calculations
      if (totalDistance > 0) {
        distances.push(totalDistance);
      }
      if (totalEmission > 0) {
        co2Emissions.push(totalEmission);
      }
    }
  
    // Sort distances and CO2 emissions for quantile and median calculations
    distances.sort((a, b) => a - b);
    co2Emissions.sort((a, b) => a - b);
  
    // Helper function to calculate quantiles
    function calculateQuantile(arr, q) {
      const pos = (arr.length - 1) * q;
      const base = Math.floor(pos);
      const rest = pos - base;
      if (arr[base + 1] !== undefined) {
        return arr[base] + rest * (arr[base + 1] - arr[base]);
      } else {
        return arr[base];
      }
    }
  
    // Helper function to calculate median
    function calculateMedian(arr) {
      return calculateQuantile(arr, 0.5);
    }
  
    // Calculate quantiles for distances
    const minDistance = distances[0] || 0;
    const q1Distance = calculateQuantile(distances, 0.25) || 0;
    const medianDistance = calculateMedian(distances) || 0;
    const q3Distance = calculateQuantile(distances, 0.75) || 0;
    const maxDistance = distances[distances.length - 1] || 0;
  
    // Calculate quantiles for CO2 emissions
    const minCo2 = co2Emissions[0] || 0;
    const q1Co2 = calculateQuantile(co2Emissions, 0.25) || 0;
    const medianCo2 = calculateMedian(co2Emissions) || 0;
    const q3Co2 = calculateQuantile(co2Emissions, 0.75) || 0;
    const maxCo2 = co2Emissions[co2Emissions.length - 1] || 0;
  
    // Prepare the data to be written to Monthly Record
    const recordRow = [
      monthLabel,       // Column A - Month (or "All - Month")
      totalPpl,         // Column B - Total People
      sumMem,           // Column C - Total BNI Members rows
      sumSub,           // Column D - Total Substitutes rows
      sumVis,           // Column E - Total Visitors rows
      minDistance,      // Column F - Minimum Distance
      q1Distance,       // Column G - First Quartile of Distance
      medianDistance,   // Column H - Median Distance
      q3Distance,       // Column I - Third Quartile of Distance
      maxDistance,      // Column J - Maximum Distance
      minCo2,           // Column K - Minimum CO2 Emissions
      q1Co2,            // Column L - First Quartile of CO2 Emissions
      medianCo2,        // Column M - Median CO2 Emissions
      q3Co2,            // Column N - Third Quartile of CO2 Emissions
      maxCo2            // Column O - Maximum CO2 Emissions
    ];
  
    // Append the new record to Monthly Record sheet
    monthlyRecordSheet.appendRow(recordRow);
    Logger.log(`Record for ${monthLabel} has been added successfully.`);
  }