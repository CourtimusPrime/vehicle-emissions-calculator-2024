const typeformMap = {
    firstName: 0,
    lastName: 1,
    emailAddress: 2,
    substituteStatus: 3,
    subbingForFirstName: 4,
    subbingForLastName: 5,
    subbingTotalDays: 6,
    joinDate: 7,
    startLocation: 8,
    stopStatus: 9,
    stopLocation: 10,
    transportMode: 11,
    transitType: 12,
    carpoolStatus: 13,
    carpoolDriverStatus: 14,
    carpoolDriverFirstName: 15,
    carpoolDriverLastName: 16,
    carDriver: 17,
    carBrand: 18,
    carModel: 19,
    carFuel: 20,
    finalLocation: 21,
    finalLocationElse: 22,
    carpoolAfterStatus: 23,
    finalTransportMode: 24,
    finalCarBrand: 25,
    finalCarModel: 26,
    finalCarFuel: 27,
    bicycleVariable: 28,
    carpoolVariable: 29,
    score: 30,
    fourthStop: 31,
    secondStop: 34,
    type: 35,
    thirdStop: 37,
    visitDate: 38,
    sixthStop: 39,
    fifthStop: 41,
    seventhStop: 42,
    eigthStop: 46,
    submitTime: 47,
    token: 48,
    username: 49
  }
  
  const bniMap = {
    token: 0,
    emailAddress: 1,
    joinDate: 2,
    startLocation: 3,
    firstStopLocation: 4,
    secondStopLocation: 5,
    thirdStopLocation: 6,
    fourthStopLocation: 7,
    transportMode: 8,
    carpoolDriverUsername: 9,
    carBrand: 10,
    carModel: 11,
    carFuel: 12,
    finalLocation: 13,
    finalTransMode: 14,
    finalCarBrand: 15,
    finalCarModel: 16,
    finalCarFuel: 17,
    fifthStopLocation: 18,
    sixthStopLocation: 19,
    seventhStopLocation: 20,
    eigthStopLocation: 21
  };
  
  const subMap = {
    token: 0,
    emailAddress: 1,
    subForToken: 2,
    subbingDays: 3,
    startLocation: 4,
    firstStopLocation: 5,
    secondStopLocation: 6,
    thirdStopLocation: 7,
    fourthStopLocation: 8,
    transportMode: 9,
    carpoolDriverUsername: 10,
    carBrand: 11,
    carModel: 12,
    carFuel: 13,
    finalLocation: 14,
    finalTransMode: 15,
    finalCarBrand: 16,
    finalCarModel: 17,
    finalCarFuel: 18,
    fifthStoplocation: 19,
    sixthStopLocation: 20,
    seventhStopLocation: 21,
    eigthStopLocation: 22
  };
  
  const visitorMap = {
    token: 0,
    emailAddress: 1,
    substituteDay: 2,
    startLocation: 3,
    firstStopLocation: 4,
    secondStopLocation: 5,
    thirdStopLocation: 6,
    fourthStopLocation: 7,
    transportMode: 8,
    carpoolDriverUsername: 9,
    carBrand: 10,
    carModel: 11,
    carFuel: 12,
    finalLocation: 13,
    finalTransMode: 14,
    finalCarBrand: 15,
    finalCarModel: 16,
    finalCarFuel: 17,
    fifthStopLocation: 18,
    sixthStopLocation: 19,
    seventhStopLocation: 20,
    eigthStopLocation: 21
  };
  
  const distanceCalculator = {
    token: 0,
    startLocation: 1,
    startTravMode: 2,
    firstStopLocation: 3,
    secondStopLocation: 4,
    thirdStopLocation: 5,
    fourthStopLocation: 6,
    fifthStopLocation: 7,
    sixthStopLocation: 8,
    seventhStoplocation: 9,
    eigthStopLocation: 10,
    distanceToFirstStop: 11,
    distanceToSecondStop: 12,
    distanceToThirdStop: 13,
    distanceToFourthStop: 14,
    distanceToConference: 15,
    conferenceLocation: 16,
    finalLocation: 17,
    finTravMode: 18,
    distanceToFifthStop: 19,
    distanceToSixthStop: 20,
    distanceToSeventhStop: 21,
    distanceToEigthStop: 22,
    distanceToFinalLocation: 23,
    organicDistance: 24,
    firstDistance: 25,
    secondDistance: 26,
    totalDistance: 27
  }
  
  const emissionsMap = {
    token: 0,           
    startDistance: 1,           
    firstCarFuel: 2,       
    secondDistance: 3,      
    secondCarFuel: 4,     
    firstEmission: 5,
    secondEmissions: 6,  
    totalEmissions: 7  
  };
  
  const carpoolMap = {
    driverToken: 0,
    partialPassengerTokens: 1, // Array
    fullPassengerTokens: 2, // Array
    driverFirstCo2: 3, // Numeric
    driverSecondCo2: 4, // Numeric
    sharedFirst: 5, // Numeric
    sharedSecond: 6 // Numeric
  }
  
  const dataSinkMap = {
    token: 0,
    type: 1,
    firstEmission: 2,
    secondEmission: 3,
    totalEmission: 4,
    totalDistance: 5
  }
  
  const recordMap = {
    month: 0,
    sumPpl: 1,
    sumMem: 2,
    sumSub: 3,
    sumVis: 4,
    minDis: 5,
    q1Dis: 6,
    medDis: 7,
    q3Dis: 8,
    maxDis: 9,
    minCo2: 10,
    q1Co2: 11,
    medCo2: 12,
    q3Co2: 13,
    maxCo2: 14
  }
  
  const archiveMap = {
    dateHandled: 0,
    timeHandled: 1,
    fullName: 2,
    token: 3,
    dateSubmitted: 4,
    type: 5,
    totalEmission: 6,
    totalDistance: 7
  }