/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the spreadsheet UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Printing')
    .addItem('PrintCut Estimate', 'showEstimatorSidebar')
    .addToUi();
}

/**
 * Creates and displays the HTML user interface in a sidebar.
 */
function showEstimatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Print Cut Estimator');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches data from the 'Print Material Inventory' sheet.
 * Returns the necessary columns for the new calculations.
 */
function getMaterialInventory() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Print Material Inventory');
    if (!sheet) {
      throw new Error("Sheet 'Print Material Inventory' not found. Please create it.");
    }
    // Fetches A2:F to get all necessary data: Name, Type, Width, Height, Cost/Sheet, Cost/LinFt
    return sheet.getRange("A2:F" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");
  } catch (e) {
    Logger.log(e);
    throw new Error(e.message);
  }
}

/**
 * Adds the final estimate values to the currently active row.
 * @param {number} qty The quantity.
 * @param {number} unitPrice The calculated price per unit.
 * @param {number} totalPrice The calculated total price.
 */
function addEstimateToProject(qty, unitPrice, totalPrice) {
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    throw new Error("No active cell selected. Please select a cell to add the project to.");
  }
  const row = activeRange.getRow();
  const sheet = activeRange.getSheet();
  
  // Set values in columns B, C, and D of the active row.
  sheet.getRange(row, 2).setValue(qty);
  sheet.getRange(row, 3).setValue(unitPrice).setNumberFormat("$#,##0.00");
  sheet.getRange(row, 4).setValue(totalPrice).setNumberFormat("$#,##0.00");
}
