function generateOPEXModel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("Setup") || ss.insertSheet("Setup");
  var costsSheet = ss.getSheetByName("Costs") || ss.insertSheet("Costs");

  // Setup Key Inputs (Adjustable by the User)
  setupSheet.clear();
  setupSheet.getRange('A1:B1').setValues([['Key Input', 'Value']]);
  setupSheet.getRange('A2:B7').setValues([
    ['Developer Hourly Rate ($)', 30],
    ['Hours per Task', 9],
    ['GCloud Cost per Developer ($)', 350],
    ['SaaS Cost per Developer ($)', 100],
    ['Office Cost for 5 Devs ($)', 500],
    ['Number of Developers', 5]
  ]);

  // Define Epics, Tasks, and Timelines
  var epics = {
    'Core Platform Stability and Security': {start: 4, end: 5, tasks: 4},
    'Model Deployment and Performance Optimization': {start: 4, end: 6, tasks: 5},
    'Tariff and Revenue Management': {start: 4, end: 5, tasks: 3},
    'User Interface and Experience': {start: 4, end: 6, tasks: 4},
    'Data Management and API Integration': {start: 7, end: 8, tasks: 5},
    'Advanced Functionality and Optimization': {start: 7, end: 8, tasks: 3}
  };

  // Prepare Costs Sheet
  costsSheet.clear();
  var headers = ['Cost Item'];
  var months = ['April 2024', 'May 2024', 'June 2024', 'July 2024', 'August 2024'];
  headers = headers.concat(months);
  costsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Fetch Key Inputs from Setup Sheet
  var values = setupSheet.getRange('B2:B7').getValues();
  var hourlyRate = values[0][0];
  var hoursPerTask = values[1][0];
  var gcloudCost = values[2][0];
  var saasCost = values[3][0];
  var officeCost = values[4][0];
  var numberOfDevelopers = values[5][0];

  var costItems = ['Developer Costs ($)', 'GCloud Servers ($)', 'SaaS Tools ($)', 'Office Space ($)', 'Total Monthly OPEX ($)'];
  var rowData = [];

  // Initialize rows for each cost item
  for (var i = 0; i < costItems.length; i++) {
    var row = new Array(months.length + 1).fill(0);
    row[0] = costItems[i]; // Set the cost item name
    rowData.push(row);
  }

  // Calculate and populate costs per month
  months.forEach(function(month, index) {
    var developerHours = 0;
    Object.keys(epics).forEach(function(epicName) {
      var epic = epics[epicName];
      if (index + 4 >= epic.start && index + 4 <= epic.end) { // Month index offset by starting month (April = 4)
        var monthsDuration = epic.end - epic.start + 1;
        var tasksPerMonth = epic.tasks / monthsDuration;
        developerHours += tasksPerMonth * hoursPerTask;
      }
    });

    var developerCosts = developerHours * hourlyRate * numberOfDevelopers;
    var gcloudServers = gcloudCost * numberOfDevelopers;
    var saasTools = saasCost * numberOfDevelopers;
    var totalOPEX = developerCosts + gcloudServers + saasTools + officeCost;

    rowData[0][index + 1] = developerCosts;
    rowData[1][index + 1] = gcloudServers;
    rowData[2][index + 1] = saasTools;
    rowData[3][index + 1] = officeCost;
    rowData[4][index + 1] = totalOPEX;
  });

  // Write costs to sheet
  costsSheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
}
