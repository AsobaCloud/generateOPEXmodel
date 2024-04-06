function generateCorrectedOPEXModel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("Setup");
  var costsSheet = ss.getSheetByName("Costs");

  if (!setupSheet) {
    setupSheet = ss.insertSheet("Setup");
    setupSheet.getRange('A1:B1').setValues([['Key Input', 'Value']]);
    setupSheet.getRange('A2:B6').setValues([
      ['Developer Hourly Rate ($)', 30],
      ['Hours per Task', 9],
      ['GCloud Cost per Developer ($)', 350],
      ['SaaS Cost per Developer ($)', 100],
      ['Office Cost for 5 Devs ($)', 500]
    ]);
  }

  if (!costsSheet) {
    costsSheet = ss.insertSheet("Costs");
  }

  costsSheet.clear();
  var months = ['April 2024', 'May 2024', 'June 2024', 'July 2024', 'August 2024'];
  costsSheet.getRange(1, 2, 1, months.length).setValues([months]);
  var costItems = ['Developer Costs ($)', 'GCloud Servers ($)', 'SaaS Tools ($)', 'Office Space ($)', 'Total Monthly OPEX ($)'];
  costsSheet.getRange('A2:A6').setValues(costItems.map(item => [item]));

  months.forEach((month, i) => {
    var col = i + 2;

    // Corrected formula for Developer Costs
    costsSheet.getRange(2, col).setFormula(`=Setup!B$2*Setup!B$3*5`); // Assuming 5 developers work on tasks each month

    // GCloud Servers, corrected formula
    costsSheet.getRange(3, col).setFormula(`=Setup!B$4*5`); // Assuming 5 developers
    
    // SaaS Tools, corrected formula
    costsSheet.getRange(4, col).setFormula(`=Setup!B$5*5`); // Assuming 5 developers
    
    // Office Space, corrected formula
    costsSheet.getRange(5, col).setFormula(`=Setup!B$6`);
    
    // Total Monthly OPEX, corrected formula to sum above costs
    costsSheet.getRange(6, col).setFormula(`=SUM(B${2 + (i*0)}:B${5 + (i*0)})`);
  });

  // Applying a Total formula for each cost item across all months, corrected to ensure correct range
  costItems.forEach((_, rowIndex) => {
    costsSheet.getRange(rowIndex + 2, months.length + 2).setFormula(`=SUM(B${rowIndex + 2}:F${rowIndex + 2})`);
  });

  // Set 'Total' label for the summation column
  costsSheet.getRange(1, months.length + 2).setValue('Total');
}
