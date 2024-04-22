// use logger.log to manually check and debug
// the only visible function to user is main function (for abstraction purpose)
// heavily use the encapsulation method to avoid redundancy
// avoid using user prompt AS MUCH AS POSSIBLE!! (to reduce error from user)
// use meaningful naming conventions
// avoid using global scope for variables

function getDetails(variable, message) {

    let variablePrompt = SpreadsheetApp.getUi().prompt(message);
    variable = variablePrompt.getResponseText();

    return variable;
}


function getServiceDetails() {

    let serviceMessage = "Please enter type of service (nTB/nA/...): ";
    let regionMessage = "Please enter the region(APAC/EMEA/AMER): ";

    let serviceDetail = getDetails("service", serviceMessage);
    let regionDetail = getDetails("region", regionMessage);
  
    let details = [serviceDetail, regionDetail];

    Logger.log("details: " + details);
    return details; 
}


function getSheet() {

    let sheetUrlPrompt = SpreadsheetApp.getUi().prompt('Enter the Sheet url:');
    let sheetUrl = sheetUrlPrompt.getResponseText();
    let spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);

    // Get the first sheet of the spreadsheet
    let sheet = spreadsheet.getSheets()[0]; 
    
    if (!spreadsheet) {
      Logger.log("spreadsheet not found");
    }

    // Return the sheet
    return sheet;
}

function findColumnHeader(sheet, keyword) {

  let currentValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();

  for (var i = 0; i < currentValues.length; i++) {
    for (var j = 0; j < currentValues[i].length; j++) {
        if (currentValues[i][j] === keyword) {
          var index = j + 1;
        }
    }
  }
  return index;
}

function countFromSourceSheet(country, caseOrigin) {

    let sourceSheet = getSheet();
    let targetColumn = findColumnHeader(sourceSheet,"Country");
    let targetColumn2 = findColumnHeader(sourceSheet, "Case Origin");
    let lastRow = sourceSheet.getLastRow() - 1; // total number of rows minus the header row

    let sourceCountry = sourceSheet.getRange(2, targetColumn, lastRow, 1);
    let sourceChannel = sourceSheet.getRange(2, targetColumn2, lastRow, 1);
    let sourceCountryValues = sourceCountry.getValues();
    Logger.log("Countries from the source sheet: " + sourceCountryValues);
    let sourceChannelValues = sourceChannel.getValues();
    Logger.log("Contact channels form the source sheet: " + sourceChannelValues);

    // Initialize counter object
    let counter = {};
    country.forEach(country => {
        caseOrigin.forEach(channel => {
            counter[country + '_' + channel] = 0;
        });
    });

    for (let i = 0; i < sourceCountryValues.length; i++) {
        for (let j = 0; j < sourceCountryValues[i].length; j++) {
            let countryValue = sourceCountryValues[i][j];
            let channelValue = sourceChannelValues[i][j];

            // Increment counter
            counter[countryValue + '_' + channelValue]++;
        }
    }
    Logger.log(counter);

    //converting the counter from object into array
    let counterArray = [];
    for (const key in counter) {
        counterArray.push([key, counter[key]]);
    }

    Logger.log("Counter based on countries and contact channels in the form of 2D Array: " + counterArray);

    return counterArray;
}

// class ID consists of service and region,as the combination of the two make a unique identifier
class ID {

      constructor(service, region) {
        this.service = service;
        this.region = region;
      }

      defineCells() {

        const services = ["nTB", "nA", "DH Support Hub"];
        const regions = ["APAC", "EMEA", "AMER", "Global"];

        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        let startColumn = 6; 
        
        let endColumn = 11;

        if (this.service === services[0] && this.region === regions[0]) {
            let startRow = 72;
            let endRow = 14;
            let nTB_APAC_cells = sheet.getRange(startRow, startColumn, endRow, endColumn);
            return nTB_APAC_cells;
        }
        // add more condition for other services and/or regions
        else if (this.service === services[0] && this.region === regions[1]) {
            let startRow = 29;
            let endRow = 22;
            let nTB_EMEA_cells = sheet.getRange(startRow, startColumn, endRow, endColumn);
            return nTB_EMEA_cells;
        }
      }
}


function updateSheet(values, service, region) { 
  
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const id = new ID(service, region);
    const initialCells = id.defineCells();
    const monthMessage = "Please enter the month: ";
    const month = getDetails("month", monthMessage);

    switch (month) {
        case months[0]:
            matchCellsByMonth(initialCells, values, 0);
            break;
        case months[1]:
            matchCellsByMonth(initialCells, values, 1);
            break;
        case months[2]:
            matchCellsByMonth(initialCells, values, 2); 
            break;
        case months[3]:
            matchCellsByMonth(initialCells, values, 3); 
            break;
        case months[4]:
            matchCellsByMonth(initialCells, values, 4); 
            break;
        case months[5]:
            matchCellsByMonth(initialCells, values, 5); 
            break;
        case months[6]:
            matchCellsByMonth(initialCells, values, 6); 
            break;
        case months[7]:
            matchCellsByMonth(initialCells, values, 7); 
            break;
        case months[8]:
            matchCellsByMonth(initialCells, values, 8); 
            break;
        case months[9]:
            matchCellsByMonth(initialCells, values, 9); 
            break;
        case months[10]:
            matchCellsByMonth(initialCells, values, 10); 
            break;
        case months[11]:
            matchCellsByMonth(initialCells, values, 11); 
            break;
        default:
            Logger.log("Invalid month entered");
            break;
    }
}

function matchCellsByMonth(initialCells, values, offsetColumn) {
    const selected_cells = initialCells.offset(0, offsetColumn, initialCells.getNumRows() + 1, 1);
    Logger.log("Values in the cells before inserting the counter: " + selected_cells.getValues());

    const valuesToSet = values.map(row => [row[1]]); // Assuming values is a 2D array
    Logger.log("Values to insert (from the counter): " + valuesToSet);

    for (let i = 0; i < valuesToSet.length; i++) {
        selected_cells.getCell(i + 1, 1).setValue(valuesToSet[i][0]); // Set values to each cell
    }

    Logger.log("Values after inserting the counter: " + selected_cells.getValues());
}

function getDetailsFromServiceAndRegions(columnOffset, service, region) { 
    const id = new ID(service, region);
    const defined_cells = id.defineCells();
    const details = defined_cells.offset(0, columnOffset, defined_cells.getNumRows() + 1, 1).getValues();
    return details;
}


function main() {

    // the positions of the columns offset from the month column, negative indicates to the left direction
    const country_column_from_month = -2;
    const contactColumn_from_month = -1;  

    const details = getServiceDetails(); // refer to the getDetails and getServiceDetails functions
    const service = details[0];
    const region = details[1];


    const country = getDetailsFromServiceAndRegions(country_column_from_month, service, region); 
    const contact_channel = getDetailsFromServiceAndRegions(contactColumn_from_month, service, region);

    updateSheet(countFromSourceSheet(country, contact_channel), service, region); 
    
}


function onOpen() {

    let ui = SpreadsheetApp.getUi();

    let menu = ui.createMenu("Automatic Counter");

    //menu.addItem("Find new Work Order", "compareWO");
    menu.addItem("Automatic counter by month", "main");


    menu.addToUi();
}


// possible inheritence class to encapsulate country and contact channel


