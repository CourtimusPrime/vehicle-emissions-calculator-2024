function distanceFunction() {
    Logger.log("Starting distance calculations...");
    measureTripToConference();
    measureTripFromConference();
    calculateTotalDistance();
    Logger.log("Distance calculations complete.");
  }
  
  function mapTravelMode(mode) {
    Logger.log("Mapping travel mode: " + mode);
  
    if (typeof mode !== 'string' || !mode) {
      Logger.log("Invalid or missing mode. Defaulting to 'driving'.");
      return 'driving';
    }
  
    switch (mode.toLowerCase()) {
      case 'car':
      case 'driving':
        return 'driving';
      case 'transit':
        return 'transit';
      case 'walking':
        return 'walking';
      case 'bicycling':
        return 'bicycling';
      default:
        Logger.log("Unknown mode: " + mode + ". Defaulting to 'driving'.");
        return 'driving';
    }
  }
  
  function measureTripToConference() {
    Logger.log("Starting measureTripToConference...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Distance Calculator');
    const frontPage = ss.getSheetByName('Front Page');
    const visitorsSheet = ss.getSheetByName('Visitors');
    const meetingAddress = frontPage.getRange('H5').getValue();
    const alternateMeetingAddress = frontPage.getRange('D5').getValue();
    const visitorTokens = visitorsSheet.getRange(2, 1, visitorsSheet.getLastRow() - 1).getValues().flat();
    const data = sheet.getDataRange().getValues();
  
    for (let i = 1; i < data.length; i++) {
      const totalDistanceValue = parseFloat(data[i][distanceCalculator.totalDistance]);
      if (!isNaN(totalDistanceValue) && totalDistanceValue > 0) {
        Logger.log(`Row ${i + 1}: Total distance already calculated (${totalDistanceValue}). Skipping row.`);
        continue;
      }
  
      const token = data[i][distanceCalculator.token];
      const startLocation = data[i][distanceCalculator.startLocation];
      const startTravMode = mapTravelMode(data[i][distanceCalculator.startTravMode]);
      const stopLocations = [
        data[i][distanceCalculator.firstStopLocation],
        data[i][distanceCalculator.secondStopLocation],
        data[i][distanceCalculator.thirdStopLocation],
        data[i][distanceCalculator.fourthStopLocation]
      ].filter(location => location && location !== '');
  
      const conferenceLocation = visitorTokens.includes(token) ? alternateMeetingAddress : meetingAddress;
      Logger.log(`Row ${i + 1}: Conference location set to ${conferenceLocation}`);
  
      let firstTripDistance = 0;
      let locations = [startLocation, ...stopLocations, conferenceLocation];
      Logger.log(`Row ${i + 1}: Locations for trip: ${locations.join(' -> ')}`);
  
      sheet.getRange(i + 1, distanceCalculator.conferenceLocation + 1).setValue(conferenceLocation);
  
      for (let j = 0; j < locations.length - 1; j++) {
        const origin = locations[j];
        const destination = locations[j + 1];
        const distanceResult = getDistance(origin, destination, startTravMode);
        const distance = distanceResult.error ? 0 : distanceResult.distance;
        firstTripDistance += distance;
  
        let distanceColumnIndex;
        if (j < stopLocations.length) {
          distanceColumnIndex = distanceCalculator[`distanceTo${['First', 'Second', 'Third', 'Fourth'][j]}Stop`];
        } else {
          distanceColumnIndex = distanceCalculator.distanceToConference;
        }
        sheet.getRange(i + 1, distanceColumnIndex + 1).setValue(distance);
        Logger.log(`Row ${i + 1}: Distance from ${origin} to ${destination}: ${distance} km`);
      }
  
      sheet.getRange(i + 1, distanceCalculator.firstDistance + 1).setValue(firstTripDistance);
      Logger.log(`Row ${i + 1}: Total distance to meeting address: ${firstTripDistance} km`);
    }
  }
  
  function measureTripFromConference() {
    Logger.log("Starting measureTripFromConference...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Distance Calculator');
    const frontPage = ss.getSheetByName('Front Page');
    const visitorsSheet = ss.getSheetByName('Visitors');
    const meetingAddress = frontPage.getRange('H5').getValue();
    const alternateMeetingAddress = frontPage.getRange('D5').getValue();
    const visitorTokens = visitorsSheet.getRange(2, 1, visitorsSheet.getLastRow() - 1).getValues().flat();
    const data = sheet.getDataRange().getValues();
  
    Logger.log(`Meeting address retrieved: ${meetingAddress}`);
  
    for (let i = 1; i < data.length; i++) {
      const totalDistanceValue = parseFloat(data[i][distanceCalculator.totalDistance]);
      if (!isNaN(totalDistanceValue) && totalDistanceValue > 0) {
        Logger.log(`Row ${i + 1}: Total distance already calculated (${totalDistanceValue}). Skipping row.`);
        continue;
      }
  
      const token = data[i][distanceCalculator.token];
      const startLocation = data[i][distanceCalculator.startLocation];
      const finalLocation = data[i][distanceCalculator.finalLocation];
      let finTravMode = data[i][distanceCalculator.finTravMode];
  
      if (!startLocation || !finalLocation) {
        Logger.log(`Row ${i + 1}: Skipping row due to missing locations. Start Location: ${startLocation}, Final Location: ${finalLocation}`);
        continue;
      }
  
      finTravMode = mapTravelMode(finTravMode || data[i][distanceCalculator.startTravMode]);
      sheet.getRange(i + 1, distanceCalculator.finTravMode + 1).setValue(finTravMode);
  
      const conferenceLocation = visitorTokens.includes(token) ? alternateMeetingAddress : meetingAddress;
      Logger.log(`Row ${i + 1}: Conference location set to ${conferenceLocation}`);
  
      const stopLocations = [
        data[i][distanceCalculator.fifthStopLocation],
        data[i][distanceCalculator.sixthStopLocation],
        data[i][distanceCalculator.seventhStopLocation],
        data[i][distanceCalculator.eighthStopLocation]
      ].filter(location => location && location !== '');
  
      let secondTripDistance = 0;
      let locations = [conferenceLocation, ...stopLocations, finalLocation];
      Logger.log(`Row ${i + 1}: Locations for return trip: ${locations.join(' -> ')}`);
  
      const distanceToFinalResult = getDistance(startLocation, finalLocation, finTravMode);
      const organicDistance = distanceToFinalResult.error ? 0 : distanceToFinalResult.distance;
      sheet.getRange(i + 1, distanceCalculator.organicDistance + 1).setValue(organicDistance);
      Logger.log(`Row ${i + 1}: Organic distance (start to final location): ${organicDistance} km`);
  
      for (let j = 0; j < locations.length - 1; j++) {
        const origin = locations[j];
        const destination = locations[j + 1];
        const distanceResult = getDistance(origin, destination, finTravMode);
        const distance = distanceResult.error ? 0 : distanceResult.distance;
        secondTripDistance += distance;
  
        let distanceColumnIndex;
        if (j < stopLocations.length) {
          distanceColumnIndex = distanceCalculator[`distanceTo${['Fifth', 'Sixth', 'Seventh', 'Eighth'][j]}Stop`];
        } else {
          distanceColumnIndex = distanceCalculator.distanceToFinalLocation;
        }
        sheet.getRange(i + 1, distanceColumnIndex + 1).setValue(distance);
        Logger.log(`Row ${i + 1}: Distance from ${origin} to ${destination}: ${distance} km`);
      }
  
      sheet.getRange(i + 1, distanceCalculator.secondDistance + 1).setValue(secondTripDistance - organicDistance);
      Logger.log(`Row ${i + 1}: Adjusted total distance from meeting address: ${secondTripDistance - organicDistance} km`);
    }
  }
  
  function calculateTotalDistance() {
    Logger.log("Starting calculateTotalDistance...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Distance Calculator');
    const data = sheet.getDataRange().getValues();
  
    for (let i = 1; i < data.length; i++) {
      const firstDistance = parseFloat(data[i][distanceCalculator.firstDistance]) || 0;
      const secondDistance = parseFloat(data[i][distanceCalculator.secondDistance]) || 0;
      const totalDist = firstDistance + secondDistance;
  
      sheet.getRange(i + 1, distanceCalculator.totalDistance + 1).setValue(totalDist);
      Logger.log(`Row ${i + 1}: Total Distance (to and from meeting address): ${totalDist} km`);
    }
  }
  
  function getDistance(origin, destination, mode) {
    Logger.log(`Requesting distance between ${origin} and ${destination} using mode ${mode}`);
    const apiKey = 'AIzaSyBkXcnjVdLDFdQ-3ElQCiNHS_gCJyxHyGk';
    const url = 'https://maps.googleapis.com/maps/api/distancematrix/json?' +
                'origins=' + encodeURIComponent(origin) +
                '&destinations=' + encodeURIComponent(destination) +
                '&mode=' + encodeURIComponent(mode) +
                '&key=' + apiKey;
  
    try {
      const response = UrlFetchApp.fetch(url);
      const jsonResponse = JSON.parse(response.getContentText());
  
      if (jsonResponse.status !== "OK") {
        Logger.log(`API Error: ${jsonResponse.error_message}`);
        return { error: true, message: `API Error: ${jsonResponse.error_message}` };
      }
  
      const element = jsonResponse.rows[0].elements[0];
      if (element.status === "OK") {
        const distanceInKm = element.distance.value / 1000;
        return { distance: distanceInKm };
      } else {
        Logger.log(`No route found: ${element.status}`);
        return { error: true, message: `No route found: ${element.status}` };
      }
    } catch (e) {
      Logger.log(`Fetch Error: ${e}`);
      return { error: true, message: `Fetch Error: ${e}` };
    }
  }