Version Control EA 1.0

// Define the background color you want to apply
  const backgroundColorYellow = '#FFFF00'; // yellow
  const backgroundColorBlue = '#4a86e8';   // cornflower blue
  const backgroundColorPurple = '#b4a7d6'; // light purple 2
  const backgroundColorRed = '#f4cccc';    // light red 3
  const backgroundColorGreen = '#b6d7a8';  // light green 2
  const backgroundColorGrey = "#999999";   // dark gray 2
  const backgroundColorGrey1 = '#b7b7b7';  // dark gray 1
  const backgroundColorG = "#cccccc";      // gray

function compareWO() {
  // Get the active sheet
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ask user to prompt yesterday's sheet date
  var sheetNamePrompt2 = SpreadsheetApp.getUi().prompt('Enter yesterday sheet date');
  var sheetName2 = sheetNamePrompt2.getResponseText();

  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName2);

  // Check if the sheet exists
  if (!sheet2) {
    Logger.log('Sheet with name "' + sheetName + '" not found.');
    return;
  }  

  let columnWO = 2;

  var data1 = sheet1.getRange(1, columnWO, sheet1.getLastRow(), 1).getValues().flat();
  var data2 = sheet2.getRange(1, columnWO, sheet2.getLastRow(), 1).getValues().flat();

  // Find new IDs in sheet1 compared to sheet2
  var newWOsSheet1 = data1.filter(workOrder => !data2.includes(workOrder));

  // Display comments for new Work Order Numbers in sheet2
  for (var j = 0; j < newWOsSheet1.length; j++) {
    var commentText = "New in " + sheet1.getName();
    var cell = sheet1.getRange(1, columnWO, sheet1.getLastRow(), 1).createTextFinder(newWOsSheet1[j]).findNext();
    if (cell) {
      cell.setComment(commentText);
    }
  }
}

function copyBackgroundColorFromOtherSheet() {
  // Get the active sheet
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the column index to copy background color
  var columnToCopy = 2; // Replace with the actual column index

  // ask the user to prompt the yesterday's sheet date
  var sheetNamePrompt1 = SpreadsheetApp.getUi().prompt('Enter yesterday sheet date');
  var sheetName1 = sheetNamePrompt1.getResponseText();

  // Get the previous sheet
  var prevSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName1);

  // Check if the sheet exists
  if (!prevSheet) {
    Logger.log('Sheet with name "' + sheetName + '" not found.');
    return;
  }  

  // Get the values and background colors from the specified column in both sheets
  var currentValues = currentSheet.getRange(1, columnToCopy, currentSheet.getLastRow(), 1).getValues();
  var prevValues = prevSheet.getRange(1, columnToCopy, prevSheet.getLastRow(), 1).getValues();
  var prevBackgroundColors = prevSheet.getRange(1, columnToCopy, prevSheet.getLastRow(), 1).getBackgrounds();

  // Iterate through each row
  for (var i = 0; i < currentValues.length; i++) {

    var matchingRowInPrevSheet = prevValues.findIndex(function(prevValue) {
      return prevValue[0] === currentValues[i][0];
    });

    if (matchingRowInPrevSheet !== -1) {
      // Set the background color in the corresponding column of the current sheet based on the background color of the cell in prevSheet
      currentSheet.getRange(i + 1, columnToCopy).setBackgrounds([[prevBackgroundColors[matchingRowInPrevSheet][0]]]);
    } else {
        continue;
    }
  }
}

function findColumnHeader(keyword) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // let keyword = "checklist";

  var currentValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();

  for (var i = 0; i < currentValues.length; i++) {
    for (var j = 0; j < currentValues[i].length; j++) {
        if (currentValues[i][j] === keyword) {
          var index = j + 1;
        }
    }
  }
  return index;
}


function setBackgroundColorForRowWithExceptions() {
  // Get the active sheet
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the base column index to base background color on
  var baseColumn = 2; // reference column

  // Define the exception column index
  var exceptionColumn = findColumnHeader("checklist"); // Exception column

  // Get the values and background colors for the entire sheet
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var backgroundColors = dataRange.getBackgrounds();

  // Loop through each row
  for (var i = 0; i < values.length; i++) {
    // Get the background color of the base column for the current row
    var baseColor = backgroundColors[i][baseColumn - 1];

    // Loop through each cell in the row
    for (var j = 0; j < values[i].length; j++) {
      if (j === baseColumn - 1 && backgroundColors[i][j] === "#ffffff") {
        break;
      }
      // Check if the cell is in the exception column
      else if (j === exceptionColumn - 1 && backgroundColors[i][j] === backgroundColorRed) {
        // Set a different color for cells in the exception column that are red
        sheet.getRange(i + 1, j + 1).setBackground(backgroundColorRed); // Set background color to red
      } else {
        // Set the background color based on the base column for other cells in the row
        sheet.getRange(i + 1, j + 1).setBackground(baseColor);
      }
    }
  }
}

function hideRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startColumn = 16; // Replace with the start column index
  var endColumn = 22;   // Replace with the end column index
  var targetColors = [backgroundColorBlue, backgroundColorGrey, backgroundColorGrey1, backgroundColorG]; 

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var backgroundColors = dataRange.getBackgrounds();

  for (var i = 0; i < values.length; i++) {
    var rowShouldBeHidden = false;

    for (var j = startColumn; j <= endColumn; j++) {
      if (targetColors.includes(backgroundColors[i][j - 1])) {
        rowShouldBeHidden = true;
        break; // No need to check other columns once a matching color is found
      }
    }

    if (rowShouldBeHidden) {
      sheet.hideRows(i + 1);
    } else {
      sheet.showRows(i + 1);
    }
  }
}

function noEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  var columnToWatch = findColumnHeader("Remark"); // Replace with the actual column index
  var flexibleKeyword = "no email"; // Replace with your flexible keyword
  var targetColumn = findColumnHeader("Contact Email");

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var cellContent = values[i][columnToWatch - 1].toLowerCase(); // Convert to lowercase for case-insensitive comparison

    // Check if the cell content contains the flexible keyword
    if (cellContent.indexOf(flexibleKeyword.toLowerCase()) !== -1) {
      sheet.getRange(i + 1, targetColumn).setBackground(backgroundColorRed); // Set background color to red
    } else {
      //sheet.getRange(i + 1, columnToWatch).setBackground("#ffffff"); // Set background color to white (or any default color)
      continue;
    }
  }
}

function onOpen() {

  let ui = SpreadsheetApp.getUi();

  var menu = ui.createMenu("Custom Functions");

  menu.addItem("Find new Work Order", "compareWO");
  menu.addItem("Copy background color based on yesterday Work Order", "copyBackgroundColorFromOtherSheet");
  menu.addItem("Copy background color for entire Row", "setBackgroundColorForRowWithExceptions");
  menu.addItem("Hide Rows", "hideRows");
  menu.addItem("No e-mail address condition", "noEmail");

  menu.addToUi();
}