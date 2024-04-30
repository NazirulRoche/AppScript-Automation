
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

function getCellsRangeInfoForID(service, region) {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let dataRange = activeSheet.getDataRange();
    let values = dataRange.getValues();

    let serviceColumn = findColumnHeader(activeSheet, "Service") - 1;
    let regionColumn = findColumnHeader(activeSheet, "Region") - 1;
    Logger.log("service column in current sheet: " + serviceColumn);
    Logger.log("region column in current sheet: " + regionColumn);

    let numberOfEntities = 0;
    let firstIndex = -1;

    for (let i = 0; i < values.length; i++) {
      if (values[i][serviceColumn] === service && values[i][regionColumn] === region) {
        numberOfEntities += 1;
        if (firstIndex === -1) {
          firstIndex = i + 1; // Update firstIndex only if it hasn't been set yet
        }
      }
    }

    //Logger.log("total number of rows: " + numberOfEntities);

    let numOfRowsAndFirstIndex = [numberOfEntities, firstIndex];

    return numOfRowsAndFirstIndex;
}

class ID {

    constructor(service, region) {
      this.service = service;
      this.region = region;
    }

    getService() {
      return this.service;
    }

    setService(service) {
      this.service = service;
    }

    getRegion() {
      return this.region;
    }

    setRegion(region) {
      this.region = region;
    }

    defineCellsForServiceAndRegion(numberOfRows, firstIndexOfServiceAndRegion) {
      Logger.log("total number of rows for " + service + " and " + region + ": " + numberOfRows);
      Logger.log("the first row index for " + service + " and " + region + ": " + firstIndexOfServiceAndRegion);

      let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      let startColumn = 6;    
      let endColumn = 12;

      let definedCells = sheet.getRange(firstIndexOfServiceAndRegion, startColumn, numberOfRows, endColumn);
      Logger.log(definedCells.getValues());
      return definedCells;
    }

}

function getServiceFromUser() {
    let message = "Please enter type of service (Please follow the Services values as in the Metric Services Sheet): ";
    let servicePrompt = SpreadsheetApp.getUi().prompt(message);
    service = servicePrompt.getResponseText();

    return service;
}

function getRegionFromUser() {
    let message = "Please enter the region(APAC/EMEA/AMER): ";
    let regionPrompt = SpreadsheetApp.getUi().prompt(message);
    region = regionPrompt.getResponseText();

    return region;
}

function getSheetURLFromUser() {

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

class SubclassID extends ID {

    constructor(service, region) {
      super(service, region);
      //this.country = country;
      //this.caseOrigin = caseOrigin;
    }

    static getSheet() {
      return getSheetURLFromUser();
    }

    countFromSourceSheet(sourceSheet) {

      let countryColumn = findColumnHeader(sourceSheet,"Country");
      let caseOriginColumn = findColumnHeader(sourceSheet, "Case Origin");
      let lastRow = sourceSheet.getLastRow() - 1; // total number of rows minus the header row

      let sourceCountry = sourceSheet.getRange(2, countryColumn, lastRow, 1);
      let sourceChannel = sourceSheet.getRange(2, caseOriginColumn, lastRow, 1);
      let sourceCountryValues = sourceCountry.getValues();
      Logger.log("Countries from the source sheet: " + sourceCountryValues);
      let sourceChannelValues = sourceChannel.getValues();
      Logger.log("Contact channels form the source sheet: " + sourceChannelValues);


      let counter = {};
      for (let i = 0; i < sourceCountryValues.length; i++) {
          counter[sourceCountryValues[i] + "_" + sourceChannelValues[i]] = 0;
      };
      Logger.log(counter);

      for (let i = 0; i < sourceCountryValues.length; i++) {
          for (let j = 0; j < sourceCountryValues[i].length; j++) {
              let countryValue = sourceCountryValues[i][j];
              let channelValue = sourceChannelValues[i][j];

              // Increment counter
              counter[countryValue + '_' + channelValue]++;

              // for NaN value we can give warning/error to user
          }
      }
      Logger.log(counter);

      //converting the counter from object into array
      let counterArray = [];
      for (const key in counter) {
            counterArray.push([key, counter[key]]);
      }

      Logger.log("counter array: " + counterArray);

      return counterArray;
    }

    sortValuesBasedOnCurrentSheetData(values, country_and_caseOriginFromCurrentSheet) {
        //const values = this.countFromSourceSheet(sourceSheet);
        //const country_and_caseOriginFromCurrentSheet = returnCountryAndCaseOriginBasedOnDefinedCells(service, region);

        let sortedValues = country_and_caseOriginFromCurrentSheet.map(function(country_and_caseOrigin) {
            for (let i = 0; i < values.length; i++) {
              if (values[i][0] !== country_and_caseOrigin) {
                continue;
              }
              else {
                Logger.log(values[i][0] + " = " + values[i][1]);
                return values[i][1];
              }
            }
            return 0;
        });
        Logger.log("sorted values: " + sortedValues);
        return sortedValues;

    }

    insertValuesToCurrentSheetBasedOnMonth(month, sortedValues, numberOfRows, firstIndexOfServiceAndRegion) {
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const initialCells = super.defineCellsForServiceAndRegion(numberOfRows, firstIndexOfServiceAndRegion);
        
        switch (month) {
        case months[0]:
            this.matchCellsByMonth(initialCells, sortedValues, 0);
            break;
        case months[1]:
            this.matchCellsByMonth(initialCells, sortedValues, 1);
            break;
        case months[2]:
            this.matchCellsByMonth(initialCells, sortedValues, 2); 
            break;
        case months[3]:
            this.matchCellsByMonth(initialCells, sortedValues, 3); 
            break;
        case months[4]:
            this.matchCellsByMonth(initialCells, sortedValues, 4); 
            break;
        case months[5]:
            this.matchCellsByMonth(initialCells, sortedValues, 5); 
            break;
        case months[6]:
            this.matchCellsByMonth(initialCells, sortedValues, 6); 
            break;
        case months[7]:
            this.matchCellsByMonth(initialCells, sortedValues, 7); 
            break;
        case months[8]:
            this.matchCellsByMonth(initialCells, sortedValues, 8); 
            break;
        case months[9]:
            this.matchCellsByMonth(initialCells, sortedValues, 9); 
            break;
        case months[10]:
            this.matchCellsByMonth(initialCells, sortedValues, 10); 
            break;
        case months[11]:
            this.matchCellsByMonth(initialCells, sortedValues, 11); 
            break;
        default:
            Logger.log("Invalid month entered");
            break;
        }
    }

    matchCellsByMonth(initialCells, sortedValues, offsetColumn) {
        const selected_cells = initialCells.offset(0, offsetColumn, initialCells.getNumRows() + 1, 1);
        Logger.log("Values in the cells before inserting the counter: " + selected_cells.getValues());

        const valuesToSet = sortedValues; 
        Logger.log("Values to insert (from the counter): " + valuesToSet);

        for (let i = 0; i < valuesToSet.length; i++) {
            selected_cells.getCell(i + 1, 1).setValue(valuesToSet[i]); // Set values to each cell
        }

        Logger.log("Values after inserting the counter: " + selected_cells.getValues());
    }



}

function getMonthFromUser() {
    let message = "Please enter the month: ";
    let regionPrompt = SpreadsheetApp.getUi().prompt(message);
    region = regionPrompt.getResponseText();

    return region;
}

function returnCountryAndCaseOriginBasedOnDefinedCells(service, region, numberOfRowsAndFirstIndex) {

    const id = new ID(service, region);
    //let numberOfRowsAndFirstIndex = getCellsRangeInfoForID(service, region);
    let service_and_region_definedCells = id.defineCellsForServiceAndRegion(numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1]);
    let country_and_caseOrigin_FromCurrentSheet = service_and_region_definedCells.offset(0, -2, service_and_region_definedCells.getNumRows(), 2).getValues();

    let country_and_caseOrigin_array = [];

    for (let i = 0; i < country_and_caseOrigin_FromCurrentSheet.length; i++) {
        country_and_caseOrigin_array.push([country_and_caseOrigin_FromCurrentSheet[i][0]] + "_" + country_and_caseOrigin_FromCurrentSheet[i][1]);
    }
    Logger.log(country_and_caseOrigin_array);

    return country_and_caseOrigin_array;
}

function main() {

   const service = getServiceFromUser();
    const region = getRegionFromUser();
    const sourceSheet = SubclassID.getSheet();
    const month = getMonthFromUser();

    const numberOfRowsAndFirstIndex = getCellsRangeInfoForID(service, region);

    const country_caseOriginInCurrentSheet = returnCountryAndCaseOriginBasedOnDefinedCells(service, region, numberOfRowsAndFirstIndex);

    Logger.log("The countries and case origins in current sheet: " + country_caseOriginInCurrentSheet);

    const subclass = new SubclassID(service, region);
    const values = subclass.countFromSourceSheet(sourceSheet);
    const sortedValues = subclass.sortValuesBasedOnCurrentSheetData(values, country_caseOriginInCurrentSheet);
    
    return subclass.insertValuesToCurrentSheetBasedOnMonth(month, sortedValues, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1]);
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();

    let menu = ui.createMenu("Automatic Counter");

    menu.addItem("Automatic counter by month", "main");


    menu.addToUi();
}


