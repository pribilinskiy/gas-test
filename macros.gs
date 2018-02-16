// https://github.com/oshliaer/lodashgs
var _ = LodashGS.load();

/**
* Convert meters to miles
* 
* @param {number} meters The distance in meters
* @return {number} The distance in miles
*/
function meterToMiles(meters) {
  
  if (typeof meters !== 'number') {
    return null;
  }
  
  return meters / 1000 * 0.621371;
}


/**
* A custom function that gets the driving distance between two addresses
*
* @param {string} origin The starting address
* @param {string} destination The target address
* @param {number} The distance in meters
*/
function drivingDistance(origin, destination) {
  
  var directions = getDirections_(origin, destination);
  
  return directions.routes[0].legs[0].distance.value;
}

/**
* A special function that runs when the spreadsheet is open,
* used to add a custom menu to the spreadhseet
*/
function onOpen() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  var menuItems = [
    {
      name: 'Generate step by step...',
      functionName: 'generateStepByStep_',
    }
  ];
  
  spreadsheet.addMenu('Directions', menuItems);
}

/**
* Creates a new sheet containing step-by-step directions
* between two user selected addresses on the "Settings" sheet
*/
function generateStepByStep() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings')
  
  settingsSheet.activate();
  
  // Prompt the user for a row number
  var selectedRow = Browser.inputBox(
    'Generate step-by-step', 
    'Please enter the row number of the addresses to use (e.g. "2"):', 
    Browser.Buttons.OK_CANCEL
  );
  
  if (selectedRow === 'cancel') {
    return
  }
  
  var rowNumber = Number(selectedRow);
  
  if (isNaN(rowNumber)
      || rowNumber < 2
      || rowNumber > settingsSheet.getLastRow())
  {
    Browser.msgBox(
      'Error', 
      Utilities.formatString('Row "%s" is not valid', selectedRow),
      Browser.Buttons.OK
    );
    return;
  }
  
  // Retrieve the addresses in that row
  var row = settingsSheet.getRange(rowNumber, 1, 1, 2);
  var rowValues = row.getValues();
  var origin = rowValues[0][0];
  var destination = rowValues[0][1];
  
  if (!origin || !destination) {
    Browser.msgBox('Error', 'Row does not contain two addresses', Browser.Buttons.OK);
    return;
  }
  
  // Get the raw directions information
  var directions = getDirections_(origin, destination);
  
  // Create a new sheet and append the steps in the directions
  var sheetName = 'Driving directions for row ' + rowNumber;
  var directionsSheet = spreadsheet.getSheetByName(sheetName);
  
  if (directionsSheet) {
    directionsSheet.clear();
    directionsSheet.activate();
  } else {
    directionsSheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  
  var sheetTitle = Utilities.formatString('Driving directions from %s to %s', origin, destination);
  
  var headers = [
    [sheetTitle, '', ''],
    ['Step', 'Distance (Meters)', 'Distance (Miles)']
  ];
  
  var newRows = [];  
  
  directions.routes[0].legs[0].steps.forEach(function(step) {
    
    // Remove HTML tags from the instruction    
    var instructions = step.html_instructions
    .replace(/<br>|<div.*?>/g, '\n')
    .replace(/<.*?>/g, '')
    ;
    
    newRows.push([
      instructions,
      step.distance.value
    ]);
    
  });
  
  directionsSheet
  .getRange(
    1, 
    1, 
    headers.length, 
    headers[0].length
  )
  .setValues(headers);
  
  directionsSheet
  .getRange(
    headers.length + 1, 
    1, 
    newRows.length, 
    newRows[0].length
  )
  .setValues(newRows);
  
  directionsSheet
  .getRange(
    headers.length + 1, 
    newRows[0].length + 1, 
    newRows.length,
    1
  )
  .setFormulaR1C1('=METERSTOMILES(R[0]C[-1])');
  
  // Format the new sheet
  directionsSheet.getRange('A1:C1').merge().setBackground('#dde');
  directionsSheet.getRange('A1:2').setFontWeight('bold');
  directionsSheet.setColumnWidth(1, 500);
  directionsSheet.getRange('B2:C').setVerticalAlignment('top');
  directionsSheet.getRange('C2:C').setNumberFormat('0.00');
  
  var stepsRange = directionsSheet.getDataRange().offset(2, 0, directionsSheet.getLastRow() - 2);
  
  setAlternatingRowBackgroundColors_(stepsRange, '#fff', '#eee');
  directionsSheet.setFrozenRows(2);
  SpreadsheetApp.flush();
}

/**
 * Sets the background colors for alternating rows within the range.
 * @param {Range} range The range to change the background colors of.
 * @param {string} oddColor The color to apply to odd rows (relative to the
 *     start of the range).
 * @param {string} evenColor The color to apply to even rows (relative to the
 *     start of the range).
 */
function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
  
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

/**
 * A shared helper function used to obtain the full set of directions
 * information between two addresses. Uses the Apps Script Maps Service.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Object} The directions response object.
 */
function getDirections_(origin, destination) {
  
  var directionFinder = Maps.newDirectionFinder();
  
  directionFinder.setOrigin(origin);
  directionFinder.setDestination(destination);
  
  var directions = directionFinder.getDirections();
  
  if (directions.routes.length == 0) {
    throw 'Unable to calculate directions between these addresses.';
  }
  
  return directions;
}
