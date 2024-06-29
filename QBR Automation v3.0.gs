// QBR Automation v3.0

// THIS SECTION IS FOR GOOGLE DRIVE, will reorganize these functions inside a class called DriveHandler
function listFilesInFolder() {
    var folderId = "1-8PKTNqaXiV_GNGK8UHZwa1YX-wlC31v"; // folder name: Service Metrics Reports (by Nazirul-Intern)
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    var filesDetails = {
          fileNames: [],
          fileIds: [],
    }

    while (files.hasNext()) {
      var file = files.next();
      //Logger.log('File Name: ' + file.getName());
      filesDetails.fileNames.push(file.getName());
      //Logger.log('File ID: ' + file.getId());
      filesDetails.fileIds.push(file.getId());
    }

    Logger.log(filesDetails.fileNames);
    Logger.log(filesDetails.fileIds);
    return filesDetails;
}

function getSheetIdFromFileId(fileIds) {
    let totalFiles = fileIds.length;
    let spreadsheetId = [];
    let file = [];
    for (let i = 0; i < totalFiles; i++) {
        file.push(DriveApp.getFileById(fileIds[i]));
        var url = [file[i].getUrl()];
        var match = /spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(url);//This part of the pattern is a capturing group enclosed in parentheses
        if (match && match[1]) {
            spreadsheetId.push(match[1]);
        } else {
            continue; // No spreadsheet ID found
        }
    }
    Logger.log("spreadsheet IDs: " + spreadsheetId);
    return spreadsheetId; // Return the matched spreadsheet ID
}

// for now this function is useless
function renameSheetNameFollowingSheetTitle(sheetNames, spreadsheetIdArray) {
    for (var i = 0; i < spreadsheetIdArray.length; i++) {
        let spreadsheet = SpreadsheetApp.openById(spreadsheetIdArray[i]);
        let sheet = spreadsheet.getSheets()[0];
        sheet.setName(sheetNames[i]);
    }

    Logger.log(sheetNames);
}


function getDetailsFromFileName(fileNames) {
    let totalFiles = fileNames.length;
    const allRegions = ["APAC", "EMEA", "AMER"];
    let service = [];
    let region = [];
    let month = [];
    let year = [];
    for (let i = 0; i < totalFiles; i++) {
        let str = fileNames[i].split("_");
        service.push(str[0]);
        if (str[1] !== "Global") {
          region.push(str[1]);
        } else { // if region is global, insert array of regions into the array
          region.push(allRegions);
        }
        month.push(str[2]);
        year.push(str[3]);
    }
    //Logger.log(service);
    //Logger.log(month);

    return {
      service,
      region,
      month,
      year
    } 
}

//SOLUTION: create another sheet that records the service along with its abbreviation
function matchServicesFromFilesWithSheet(fileNames) {

    const servicesFromFilesName = getDetailsFromFileName(fileNames).service;

    const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Services Codes");

    const serviceCodes = referenceSheet.getRange(2, 1, referenceSheet.getLastRow() - 1, 2).getValues();
    
    const serviceObject = {};
    for (let i  = 0; i < serviceCodes.length; i++) {
        serviceObject[serviceCodes[i][1]] = serviceCodes[i][0];
    }

    const serviceKeys = Object.keys(serviceObject);

    let services = [];

    servicesFromFilesName.forEach(function(service) {
        if (serviceKeys.includes(service)) {
            services.push(serviceObject[service]);
        }
    });

    //Logger.log("services: " + services);
    return services;     
}

// THIS SECTION IS FOR GOOGLE SHEET
//TODO: Change into class called SheetHandler
// monthNames can be declared in the Main script globally
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

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

// ISSUE: if there are more than 1 column with keyword "opened" there will be error
function findOpenedDateColumn(referenceSheet) {
    const columnHeader = referenceSheet.getRange(1, 1, 1, referenceSheet.getLastColumn()).getValues().flat();
    const keyword = "Opened";
    for (let i = 0; i < columnHeader.length; i++) {
        if (columnHeader[i].includes(keyword.toLowerCase())) var index = i + 1;
    }
    return index;
}


function getCellsRangeInfoForID(service, region) {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let dataRange = activeSheet.getDataRange();
    let values = dataRange.getValues();

    let serviceColumn = findColumnHeader(activeSheet, "Service") - 1;
    let regionColumn = findColumnHeader(activeSheet, "Region") - 1;

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

    let numOfRowsAndFirstIndex = [numberOfEntities, firstIndex];

    return numOfRowsAndFirstIndex;
}

// PURPOSE: Standardize the format of date data into months (data cleaning/processing)
function getDateInfo(referenceSheet, keyword, currentYear) {
    let column = findColumnHeader(referenceSheet, keyword);
    let values = referenceSheet.getRange(2, column, referenceSheet.getLastRow(), 1).getDisplayValues();
    const totalValues = values.length - 1; // column header is not included
    let monthsFromReferenceSheet = [];

    //Logger.log(values);
    let years = [];

    for (let i = 0; i < totalValues; i++) {
        var parts = values[i][0].split('/');
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1; // Subtract 1 because months are zero-indexed in JavaScript
        var year = parseInt(parts[2], 10);
        var date = new Date(year, month, day);

        years.push(year);
        
        if (date.getFullYear() === currentYear) monthsFromReferenceSheet.push(monthNames[date.getMonth()]);
    }

    //Logger.log(monthsFromReferenceSheet);
    return {
        monthsFromReferenceSheet,
        years
    };
}

// this function may be obsolete in the future, since months and years are based on the reports' title
function getCurrentTime() {
  // Create a new Date object for the current date and time
  var currentDate = new Date();
  
  // Get the numeric month (0 for January, 1 for February, etc.) and add 1 to make it human-readable
  var numericMonth = currentDate.getMonth() + 1;
  var currentYear = currentDate.getFullYear();

  for (let i = 1; i < monthNames.length + 1; i++) {
      if (numericMonth === i) var monthName = monthNames[i - 1]; // minus 1 because monthNames index start with 0
  }

  // Return an object containing both representations
  return {
    monthName: monthName,
    currentYear: currentYear
  };
}


class ID {
    //declare as object so that the order that does not matter
    constructor({service, region, month}) {
      this.service = service;
      this.region = region || null;
      this.month = month || null;
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

    getMonth() {
      return this.month;
    }

    setMonth(month) {
      this.month = month;
    }

    defineCells(numberOfRows, firstIndexOfServiceAndRegion, month) {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      let startColumn = 6;    
      let endColumn = 12;

      let definedCellsForServiceAndRegion = sheet.getRange(firstIndexOfServiceAndRegion, startColumn, numberOfRows, endColumn);
      var offsetColumn;

      for (let i = 0; i < monthNames.length; i++) {
          offsetColumn = i;
          if (month === monthNames[i]) {
              var definedCells = definedCellsForServiceAndRegion.offset(0, offsetColumn, numberOfRows, 1);
          }
      }
      
      return definedCells;
    }

}

// nTB main process is done, add-on to count Obselete cases (also done)
// For reports that are similar to nTB
// Its composite key is combination of service, region, month, country and case origin
// Since country and case origin can directly be accessed from the reference sheet, its not included in the constructor
class ProcessOne extends ID {

    countFromReferenceSheet(referenceSheet) {
      let countryColumn = findColumnHeader(referenceSheet,"Country");
      let caseOriginColumn = findColumnHeader(referenceSheet, "Case Origin");
      let statusColumn = findColumnHeader(referenceSheet, "Status");
      let lastRow = referenceSheet.getLastRow() - 1; // total number of rows minus the header row

      let sourceCountry = referenceSheet.getRange(2, countryColumn, lastRow, 1);
      let sourceChannel = referenceSheet.getRange(2, caseOriginColumn, lastRow, 1);
      let status = referenceSheet.getRange(2, statusColumn, lastRow, 1);

      let sourceCountryValues = sourceCountry.getValues();
      let sourceChannelValues = sourceChannel.getValues();
      let statusValues = status.getValues().flat();
      Logger.log(statusValues);

      let countObsoleteCase = {};

      let counter = {};
      for (let i = 0; i < sourceCountryValues.length; i++) {
          counter[sourceCountryValues[i] + "_" + sourceChannelValues[i]] = 0;    
          countObsoleteCase[sourceCountryValues[i] + "_" + sourceChannelValues[i]] = 0; 
      };
      
      const keyword = "Obsolete";

      for (let i = 0; i < sourceCountryValues.length; i++) {
          let countryValue = sourceCountryValues[i][0];
          let channelValue = sourceChannelValues[i][0];
          counter[countryValue + '_' + channelValue]++;

          if (statusValues[i] === keyword) countObsoleteCase[countryValue + "_" + channelValue]++;
      }
      
      Logger.log(counter);
      Logger.log(countObsoleteCase);

      return {
          counter,
          countObsoleteCase
      };
    }     

    //dependent to getCellsRangeInfoForID (to access numberOfRows and firstIndexOfServiceAndRegion)
    // dependent to countFromReferenceSheet
    sortValuesBasedOnCurrentSheetData(values, numberOfRows, firstIndexOfServiceAndRegion) {
        let mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const countryColumn = findColumnHeader(mainSheet, "Country");
        const caseOriginColumn = findColumnHeader(mainSheet, "Contact Channel");
        const countries = mainSheet.getRange(firstIndexOfServiceAndRegion, countryColumn, numberOfRows, 1).getValues();
        const caseOrigins = mainSheet.getRange(firstIndexOfServiceAndRegion, caseOriginColumn, numberOfRows, 1).getValues();

        let countryAndCaseOriginArray = [];

        for (let i = 0; i < countries.length; i++) {
            countryAndCaseOriginArray.push(countries[i] + "_" + caseOrigins[i]);
        }

        //Logger.log(countryAndCaseOriginArray);
        
        let keys = Object.keys(values);
        Logger.log("values: " + keys);

        let sortedValues = [];

        for (let i = 0; i < countryAndCaseOriginArray.length; i++) {
            if (keys.includes(countryAndCaseOriginArray[i])) {
              sortedValues.push(values[countryAndCaseOriginArray[i]]); 
            }
            else {
              sortedValues.push(0); 
            }
        }

        Logger.log("sorted values: " + sortedValues);
        return sortedValues;
    }

    insertValuesIntoMainSheet(sortedValues, numberOfRows, firstIndexOfServiceAndRegion, month) {
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, month);
        for (let i = 0; i < sortedValues.length; i++) {
            definedCells.getCell(i + 1, 1).setValue(sortedValues[i]);
        }
    }

    isThereObseleteCases(obseleteCasesCount, numberOfRows, firstIndexOfServiceAndRegion, month) {
        let keys = Object.keys(obseleteCasesCount);
        let values = Object.values(obseleteCasesCount);

        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const countryColumn = findColumnHeader(mainSheet, "Country");
        const caseOriginColumn = findColumnHeader(mainSheet, "Contact Channel");

        let cellsRange = mainSheet.getRange(firstIndexOfServiceAndRegion, 1, numberOfRows, mainSheet.getLastColumn());
        let cellsValues = cellsRange.getValues();
        
        let obseleteCasesArray = [];
        let countryAndCaseOrigin = [];

        for (let i = 0; i < values.length; i++) {
            if (values[i] !== 0) {
                countryAndCaseOrigin.push(keys[i]);
                obseleteCasesArray.push(values[i]);
            }
        }

        if (obseleteCasesArray.length === 0) {
            return 0;
        }
        else {
            let indexArray = [];
            let countryAndCaseOriginObject = {}
            for (let i = 0; i < obseleteCasesArray.length; i++) {
                countryAndCaseOriginObject[countryAndCaseOrigin[i].split("_")[0]] = countryAndCaseOrigin[i].split("_")[1];
            }

            Logger.log(countryAndCaseOriginObject);
            obseleteCasesArray = [];
            
            for (let i = 0; i < cellsValues.length; i++) {
                for (let key in countryAndCaseOriginObject) {
                    if (cellsValues[i][countryColumn - 1] === key && cellsValues[i][caseOriginColumn - 1] === countryAndCaseOriginObject[key]) {
                        let rowIndex = cellsRange.getCell(i + 1, countryColumn).getRow();
                        indexArray.push(rowIndex);
                    }
                }
                for (let key in obseleteCasesCount) {
                    //Logger.log(cellsValues[i][countryColumn - 1] + "_" + cellsValues[i][caseOriginColumn - 1]);
                    //Logger.log(key);
                    if ((cellsValues[i][countryColumn - 1] + "_" + cellsValues[i][caseOriginColumn - 1]) === key && obseleteCasesCount[key] !== 0) obseleteCasesArray.push(obseleteCasesCount[key]);
                }
            }
            Logger.log("obsolete cases array: " + obseleteCasesArray);
            Logger.log(indexArray);

            let definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, month);
        
            for (let i = 0; i < indexArray.length; i++) {
                let commentText = "Obselete cases count: " + obseleteCasesArray[i];
                let cell = definedCells.offset(indexArray[i] - firstIndexOfServiceAndRegion, 0, 1);
                cell.setComment(commentText);
            }
          }
      }

}

// for Report Structure similar to Navify Analytics APAC
// composite key is combination of service, region and month
// involves current year so if there is transition of year, it may not work (instead of currentYear, try to use reports' year)
class ProcessTwo extends ID {

    insertValuesIntoCells(referenceSheet, keyword, numberOfRows, firstIndexOfServiceAndRegion, currentYear, selectedMonth) {
        const { monthsFromReferenceSheet, years } = getDateInfo(referenceSheet, keyword, currentYear);

        const statusColumn = findColumnHeader(referenceSheet, "Status");

        let statusValues = referenceSheet.getRange(2, statusColumn, referenceSheet.getLastRow() - 1, 1).getValues().flat();

        let statusForCurrentYear = [];
        for (let i = 0; i < years.length; i++) {
            if (years[i] === currentYear) statusForCurrentYear.push(statusValues[i]);
        }

        //Logger.log(statusForCurrentYear);

        let count = 0;

        for (let i = 0; i < monthsFromReferenceSheet.length; i++) {
            if (monthsFromReferenceSheet[i] === selectedMonth) count++;
        }

        const obsoleteKeyword = "Obsolete";

        let obseleteCount = 0;

        for (let i = 0; i < statusForCurrentYear.length; i++) {
            if (statusForCurrentYear[i] === obsoleteKeyword && monthsFromReferenceSheet[i] === selectedMonth) obseleteCount++;
        }

        Logger.log("obsolete count: " + obseleteCount);
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
    
        definedCells.setValue(count);
        Logger.log("defined cells notation: " + definedCells.getA1Notation());

        if (obseleteCount === 0) {
            return 0;
        }
        else {
            const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
            const commentText = "Obsolete cases count: " + obseleteCount;
            definedCells.setComment(commentText);
        }
    }

}

// composite key is combination of service, region and month
// the region in the report is global, needs to classify the city/country into region by ourself
// TODO: Discuss the classification with PO
class ProcessThree extends ID {

    classifyLocations(referenceSheet, currentYear, selectedMonth, region) {
        const locationColumn = findColumnHeader(referenceSheet, "Location");
        let locationValues = referenceSheet.getRange(2, locationColumn, referenceSheet.getLastRow() - 1, 1).getValues().flat();

        const { monthsFromReferenceSheet, years } = getDateInfo(referenceSheet, "Opened", currentYear);

        let locations = [];
        for (let i = 0; i < years.length; i++) {
            if (years[i] === currentYear) {
                locations.push(locationValues[i]);
            }
        }

        Logger.log(locations.length);
        Logger.log(locations);
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Regions Classification");

        const apacColumn = findColumnHeader(sheet, "APAC Region");
        const emeaColumn = findColumnHeader(sheet, "EMEA Region");
        const amerColumn = findColumnHeader(sheet, "AMER Region");
        const total = sheet.getLastRow() - 1; // minus header column

        let apacLocations = sheet.getRange(2, apacColumn, total, 1).getValues().flat();
        let emeaLocations = sheet.getRange(2, emeaColumn, total, 1).getValues().flat();
        let amerLocations = sheet.getRange(2, amerColumn, total, 1).getValues().flat();

        let countApac = 0;
        let countEmea = 0;
        let countAmer = 0;

        for (let i = 0; i < locations.length; i++) {
            if (apacLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countApac += 1;
            } else if (emeaLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countEmea += 1;
            } else if (amerLocations.includes(locations[i]) && monthsFromReferenceSheet[i] === selectedMonth) {
                countAmer += 1;
            }
        }

        Logger.log(countApac);
        Logger.log(countEmea);
        Logger.log(countAmer);

        if (region === "APAC") {
            return countApac;
        } else if (region === "EMEA") {
            return countEmea;
        } else if (region === "AMER") {
            return countAmer;
        } else {
            Logger.log("invalid region");
        }
    }

    insertValuesByUpdatedRegion(count, service, selectedMonth, region) {
        let numOfRowsAndFirstIndex = getCellsRangeInfoForID(service, region);
        let definedCells = super.defineCells(numOfRowsAndFirstIndex[0], numOfRowsAndFirstIndex[1], selectedMonth);
        definedCells.setValue(count);  
    }

}

// for ISIS report (this report structure is similar to nA APACS)
// date is retrieved from "opened_at" column
// generate the report on monthly basis
class ProcessFour extends ID {

    insertValuesIntoCells(referenceSheet, currentYear, numberOfRows, firstIndexOfServiceAndRegion, selectedMonth) {

        const monthsFromReferenceSheet = getDateInfo(referenceSheet, "opened_at", currentYear).monthsFromReferenceSheet;

        let count = 0;

        for (let i = 0; i < monthsFromReferenceSheet.length; i++) {
            if (monthsFromReferenceSheet[i] === selectedMonth) count += 1;
        }
        Logger.log(count);
        const definedCells = super.defineCells(numberOfRows, firstIndexOfServiceAndRegion, selectedMonth);
        definedCells.setValue(count);
    }

}


function determineReportType(service, region, month, selectedMonth) {

    const reportOne = "first type";
    const reportTwo = "second type";
    const reportThree = "third type";
    const reportFour = "fourth type";

    if (service === "Navify Tumorboard" && region === "APAC" && month === selectedMonth) {
        return reportOne;
    } else if (service === "Navify Analytics" && region === "APAC") {
        return reportTwo;
    } else if (service === "SIP") {
        return reportThree;
    } else if (service === "ISIS") {
        return reportFour;
    } else {
        return 0;
    }

}

function main(r) { // r is the name of the specific region
    const message = "Please choose a month (Jan/Feb/Mar/Apr/May/Jun/Jul/Aug/Sep/Oct/Nov/Dec): ";
    let monthPrompt = SpreadsheetApp.getUi().prompt(message);
    let selectedMonth = monthPrompt.getResponseText();

    const currentYear = getCurrentTime().currentYear;

    const { fileNames, fileIds } = listFilesInFolder();
    const services = matchServicesFromFilesWithSheet(fileNames);
    const { region: region, month: monthsFromFile } = getDetailsFromFileName(fileNames);

    for (let i = 0; i < region.length; i++) {
      if (Array.isArray(region[i])) {
          for (let j = 0; j < region[i].length; j++) {
            if (region[i][j] === r) region[i] = region[i][j];
          }
      }
    }

    // Filter file details based on selected region
    const filteredIndexes = region.reduce((acc, reg, idx) => {
        if (reg === r) {
            acc.push(idx); 
        }
        Logger.log(acc);
        return acc;
    }, []);

    Logger.log("region: " + region);

    const filteredServices = filteredIndexes.map(idx => services[idx]);
    const filteredMonths = filteredIndexes.map(idx => monthsFromFile[idx]);
    const filteredRegions = filteredIndexes.map(idx => region[idx]);
    const filteredSheetIds = filteredIndexes.map(idx => fileIds[idx]);

    Logger.log("filtered indexes: " + filteredIndexes);
    Logger.log("filtered services: " + filteredServices);
    Logger.log("filtered months: " + filteredMonths);
    Logger.log("filtered sheet IDs: " + filteredSheetIds);
    Logger.log("Sheet IDs length: " + filteredSheetIds.length);

    // Iterate over filtered data for the selected region
    for (let i = 0; i < filteredSheetIds.length; i++) {
        const referenceSpreadsheet = SpreadsheetApp.openById(filteredSheetIds[i]);
        const referenceSheet = referenceSpreadsheet.getSheets()[0];

        const numberOfRowsAndFirstIndex = getCellsRangeInfoForID(filteredServices[i], filteredRegions[i]);

        Logger.log("service: " + filteredServices[i]); 
        Logger.log("region: " + filteredRegions[i]);
        Logger.log("month: " + filteredMonths[i]);
        Logger.log("selected month: " + selectedMonth);

        let reportType = determineReportType(filteredServices[i], filteredRegions[i], filteredMonths[i], selectedMonth);

         Logger.log("report type: " + reportType);

        switch(reportType) {
            case "first type":
                let processOne = new ProcessOne({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                let {counter, countObsoleteCase} = processOne.countFromReferenceSheet(referenceSheet);
                let sortedValues = processOne.sortValuesBasedOnCurrentSheetData(counter, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1]);
                processOne.insertValuesIntoMainSheet(sortedValues, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                processOne.isThereObseleteCases(countObsoleteCase, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                break;
            case "second type":
                let processTwo = new ProcessTwo({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processTwo.insertValuesIntoCells(referenceSheet, "Date/Time Opened", numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], currentYear, selectedMonth);
                break;
            case "third type":
                Logger.log("region: " + region[i]);
                let processThree = new ProcessThree({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                const count = processThree.classifyLocations(referenceSheet, currentYear, selectedMonth, region[i]);
                processThree.insertValuesByUpdatedRegion(count, services[i], selectedMonth, region[i]);
                break;
            case "fourth type":
                let processFour = new ProcessFour({service: filteredServices[i], region: filteredRegions[i], month: filteredMonths[i]});
                processFour.insertValuesIntoCells(referenceSheet, currentYear, numberOfRowsAndFirstIndex[0], numberOfRowsAndFirstIndex[1], selectedMonth);
                break;
            default:
                Logger.log("Process is not yet established or report is not related.");
                break;
        }
    }
}

function mainAPAC() {
    main("APAC");
}

function mainEMEA() {
    main("EMEA");
}

function mainAMER() {
    main("AMER");
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();

    var menu = ui.createMenu("Automatic Count");

    menu.addItem("Count for APAC region", "mainAPAC");
    menu.addItem("Count for EMEA region", "mainEMEA");
    menu.addItem("Count for AMER region", "mainAMER");

    menu.addToUi();
}
