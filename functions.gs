function templateSheetID() {
  // Set template spreadsheet - need to wrap variables in function for onOpen function to work
  const SS = SpreadsheetApp.openById("INPUT ID")
  return SS
}


function templateSlideID() {
  // Set template spreadsheet - need to wrap variables in function for onOpen function to work
  const SS = SlidesApp.openById("INPUT ID")
  return SS
}

function projectID() {
  // Get project ID inputted
  const PID = SpreadsheetApp.getActive().getRange("Dashboard!B13").getValue();
  return PID
}


function getData(dataURL, sheetName, sourceSheet , projectID, columnIndexNo) {
  Logger.log("Start get data")

  // Get the 'data' sheet and clear existing content
  const sheet = sourceSheet.getSheetByName(sheetName);
  sheet.clearContents();

  // URL of CSV file
  const url = dataURL

  // Fetch the CSV data
  const csv = UrlFetchApp.fetch(url);
  Logger.log('Fetch data')

  // Parse CSV data
  const data = Utilities.parseCsv(csv);

  // Extract headers from the first row of the CSV data
  const headers = data.shift();

  const columnIndex = columnIndexNo
  const condition = projectID; // Filter by Project ID
  Logger.log("Project ID: " + condition)

  // Filter rows based on the condition
  const filteredData = data.filter(row => row[columnIndex] === condition);

  // Insert headers into the sheet at the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get the dimensions of the filtered data
  const numRows = filteredData.length;
  const numCols = data[0].length;

  // Get the range starting from cell A1
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Set values in the sheet
  range.setValues(filteredData);
  Logger.log("End get data")
}

function getBenchmarks(dataURL, sheetName, sourceSheet, columnIndexNo) {
  logMsg("Start getBenchmarksCsv")

  // Get the 'data' sheet and clear existing content
  var benchmarkSheet = sourceSheet.getSheetByName(sheetName);
  benchmarkSheet.clearContents();
  benchmarkSheet.clear();

  // URL of CSV file
  var url = dataURL;

  // Fetch the CSV data
  var csv = UrlFetchApp.fetch(url);

  // Parse CSV data
  var data = Utilities.parseCsv(csv);

  // Extract headers from the first row of the CSV data
  var headers = data.shift(); // Remove and store the first row as headers

  // Define sheet variables
  var projectSheet = sourceSheet.getSheetByName("Project");

  // Get all unqiue platform, region, country values
  // 1. Platform
  var nPlatforms = projectSheet.getRange("B11").getValue();
  Logger.log("N Platforms: " + nPlatforms);

  var platformRange = projectSheet.getRange(12, 2, nPlatforms, 1);
  Logger.log("Platforms: " + platformRange.getA1Notation());

  // 2. Regions
  var nRegions = projectSheet.getRange("C11").getValue();
  Logger.log("N Regions: " + nRegions);

  var regionRange = projectSheet.getRange(12, 3, nRegions, 1);
  Logger.log("Regions: " + regionRange.getA1Notation());

  // 3. Country
  var nCountry = projectSheet.getRange("D11").getValue();
  Logger.log("N Country: " + nCountry);

  var countryRange = projectSheet.getRange(12, 4, nCountry, 1);
  Logger.log("Country: " + countryRange.getA1Notation());

  // Get values
  var platform = platformRange.getValues();
  Logger.log("Platforms: " + platform);

  var region = regionRange.getValues();
  Logger.log("Regions: " + region);

  var country = countryRange.getValues();
  Logger.log("Countries: " + country);

  // Get all combinations of benchmarks
  // Platform-Region
  var benchmarks = [];

  // Platform-Region Benchmarks
  // Iterate through each platform
  for (let p = 0; p < platform.length; p++) {
    // Iterate through each region
    for (let r = 0; r < region.length; r++) {
      // Combine platform and region and add to the benchmarks array
      benchmarks.push(platform[p][0].toLowerCase() + "-" + region[r][0].toLowerCase());
    }
  }
  Logger.log("Add Platform-Region:" + benchmarks);

  // Platform-Country Benchmarks
  // Iterate through each platform
  for (let p = 0; p < platform.length; p++) {
    // Iterate through each country
    for (let c = 0; c < country.length; c++) {
      // Combine platform and country and add to the benchmarks array
      benchmarks.push(platform[p][0].toLowerCase() + "-" + country[c][0].toLowerCase());
    }
  }
  Logger.log("Add Platform-Country:" + benchmarks);

  // Platform Benchmarks
  // Iterate through each platform
  for (let p = 0; p < platform.length; p++) {
    benchmarks.push(platform[p][0].toLowerCase());
  }
  Logger.log("Add Platform:" + benchmarks);


  // Filter rows based on the condition
  var filteredData = data.filter(row => benchmarks.includes(row[columnIndexNo].toLowerCase()));

  // Insert headers into the sheet at the first row
  benchmarkSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get the dimensions of the filtered data
  var numRows = filteredData.length;
  var numCols = data[0].length;

  // Get the range starting from cell A1
  var range = benchmarkSheet.getRange(2, 1, numRows, numCols);

  // Set values in the sheet
  range.setValues(filteredData);

  logMsg("End getBenchmarksCsv")

}

function createNamedRange(sheetName, headerNamedRange, dataNamedRange) {
  // Define sheet
  var sheet = templateSheetID().getSheetByName(sheetName);
  Logger.log(sheetName)

  var lastCol = sheet.getRange(1, sheet.getLastColumn()).getA1Notation(); // Get last column in A1 notation

  // Create header named range
  var headerRange = sheet.getRange("A1:" + lastCol);
  Logger.log("Header Range: " + headerRange.getA1Notation())
  templateSheetID().setNamedRange(headerNamedRange, headerRange);
  
  // Logger.log("Header Range: " + headerRange.getA1Notation())

  // Create data named range
  var dataRange = sheet.getDataRange();
  templateSheetID().setNamedRange(dataNamedRange, dataRange);

  Logger.log("Data Range: " + dataRange.getA1Notation())
}

function getBrandLogo(projectID, reportFolderID) {
  // URL of CSV file
  var url = 'http://business-insights.elementhuman.com/public/question/2e692b64-d2ae-4d73-a6a2-90fdc8ce03d0.csv';

  // Fetch the CSV data
  var csv = UrlFetchApp.fetch(url);

  // Parse CSV data
  var data = Utilities.parseCsv(csv);

  // Filter rows based on the condition
  var filteredData = data.filter(row => row[1] === projectID);
  Logger.log(filteredData)

  // Store logo link
  var logoUrl = filteredData[0][5]
  Logger.log(logoUrl)

  var logoName = filteredData[0][4]
  Logger.log(logoName)

  // Save image to folder
  var response = UrlFetchApp.fetch(logoUrl);
  var fileBlob = response.getBlob()
  var folder = DriveApp.getFolderById(reportFolderID)
  var logoFile = folder.createFile(fileBlob).setName('logo_' + logoName);
  return logoFile
}


function getStimGif(activityID, reportFolderID) {
  // URL of CSV file
  var url = 'dummyurl';

  // Fetch the CSV data
  var csv = UrlFetchApp.fetch(url);

  // Parse CSV data
  var data = Utilities.parseCsv(csv);

  // Filter rows based on the condition
  var filteredData = data.filter(row => row[4] === activityID);
  Logger.log(filteredData)

  // Store logo link
  var logoUrl = filteredData[0][7]
  Logger.log(logoUrl)

  var logoName = filteredData[0][5]
  Logger.log(logoName)

  // Save image to folder
  var response = UrlFetchApp.fetch(logoUrl);
  var fileBlob = response.getBlob()
  var folder = DriveApp.getFolderById(reportFolderID)
  var logoFile = folder.createFile(fileBlob).setName('gif_' + logoName);
  return logoFile
}

function insertBrandLogo(slide, logoFile, left, top) {
  // Insert image into slide
  const logo = slide.insertImage(logoFile).setLeft(left).setTop(top);
  const logoHeight = logo.getHeight();
  const logoWidth = logo.getWidth();
  logo.scaleHeight(37 / logoHeight).scaleWidth(114 / logoWidth);
}

function importDataSource() {
  logMsg("Import data source");
  // Spreadsheet Data Source
  Logger.log(projectID())
  // Get Activity Brand Metrics data
  getData("INPUT URL", "Activity Brand Metrics", templateSheetID(), projectID(), 6);
  // Get Activity Implicit per Trait
  getData("INPUT URL", "Activity Implicit per Trait", templateSheetID(), projectID(), 6);
  // Get Activity Expression data
  getData("INPUT URL", "Activity Expression", templateSheetID(), projectID(), 4);
  // Get Benchmarks
  getBenchmarks("INPUT URL", "Benchmarks", templateSheetID(), 1)

  // Set Named Range for each sheet
  logMsg("Set Named Range for each sheet")
  createNamedRange("Activity Brand Metrics", "activity_brand_metrics_header", "activity_brand_metrics");
  createNamedRange("Activity Implicit per Trait", "activity_implicit_trait_header", "activity_implicit_trait_metrics");
  createNamedRange("Activity Expression", "activity_expression_header", "activity_expression_metrics");
  createNamedRange("Benchmarks", "benchmarks_header", "benchmarks_data");

  // Create scorecard benchmark named range
  var sheet = templateSheetID().getSheetByName("Scorecard Benchmark");
  var lastCol = sheet.getLastColumn();
  // Create header named range
  var dataRange = sheet.getRange(1,1,1,lastCol);
  Logger.log("Benchmark header: " + dataRange.getA1Notation());
  templateSheetID().setNamedRange("scorecard_benchmark_header", dataRange);

  // Create data named range
  var dataRange = sheet.getRange(2, 1, 1, lastCol);
  Logger.log("Benchmark value: " + dataRange.getA1Notation());
  templateSheetID().setNamedRange("scorecard_benchmark_value", dataRange);

  logMsg("END import data source")
}

function calculateBenchmarkDiff(){
  logMsg("Calculate Benchmark Difference in Template Spreadsheet")
  var sheet = templateSheetID().getSheetByName("Scorecard")
  sheet.getRange("K5:AB").clear(); // clear existing content

  var nExposed = templateSheetID().getRange("Project!B4").getValue();

  if (nExposed > 1){
  var formulaRange = sheet.getRange("K4:AB4")
  var destinationRange = sheet.getRange(5, 11, nExposed-1, 18);
  formulaRange.copyTo(destinationRange);
  }
}

function createReportFolder(){
  logMsg("Create report folder and template copies")
  // Get project name
  var reportName = templateSheetID().getRange("B6").getValue()
  Logger.log(reportName);

  // Create a report folder
  // Get the parent report folder
  var parentFolder = DriveApp.getFolderById("");

  // Create a project subfolder
  var scorecardFolder = parentFolder.createFolder(reportName);

  // Make a copy of the spreadsheet template and save in the project subfolder
  var spreadsheetCopy = DriveApp.getFileById(templateSheetID().getId()).makeCopy(scorecardFolder);
  spreadsheetCopy.setName("Data_Source_" + reportName); // Set name

  // Make a copy of the slide template report and save in the project subfolder
  var slideCopy = DriveApp.getFileById(templateSlideID().getId()).makeCopy(scorecardFolder);
  slideCopy.setName("Scorecard_" + reportName); // Set name

  var dataSpreadsheet = SpreadsheetApp.openById(spreadsheetCopy.getId());
  var scorecardSlides = SlidesApp.openById(slideCopy.getId());

  Logger.log("Spreadsheet: " + spreadsheetCopy.getId());
  Logger.log("Slides: " + slideCopy.getId());

  return [dataSpreadsheet, scorecardSlides, scorecardFolder, reportName]
}

function createActivitySlides(dataSpreadsheet, scorecardSlides, scorecardFolder){
  logMsg("Create activity slides");

  // Slide 1
  logMsg("Slide 1")
  var slide = scorecardSlides.getSlides()[0]
  var slidebyID = scorecardSlides.getSlideById(slide.getObjectId())
  Logger.log(slide + " " + slidebyID)

  // Get Main Brand logo
  Logger.log("Insert brand logo");
  Logger.log(scorecardFolder.getId());
  try{
  var logoFile = getBrandLogo(projectID(), scorecardFolder.getId());

  // Insert Brand logo
  insertBrandLogo(slide, logoFile, 36, 217);
  } catch(error){
    logMsg("Unable to insert main brand logo. " + error)
  }


  // Get and insert titles
  logMsg("Get and insert titles");
  var projectName = dataSpreadsheet.getRange("Project!B3").getValue().toUpperCase();
  Logger.log(projectName);
  var mainBrand = dataSpreadsheet.getRange("Scorecard!C4").getValue().toUpperCase();
  Logger.log("Main Brand" + mainBrand);
  var date = dataSpreadsheet.getRange("Project!B5").getValue();

  slidebyID.replaceAllText('{{projectName}}',projectName);
  slidebyID.replaceAllText('{{mainBrand}}',mainBrand);
  slidebyID.replaceAllText('{{date}}',date);
  slidebyID.replaceAllText('{{projectID}}',projectID());
  

  // Slide 3+

  // Make a copy of scorecard slide based on # of exposed activities
  var nExposed = dataSpreadsheet.getRange("Project!B4").getValue()
  Logger.log("# of Exposed: " + nExposed);

  var slide = scorecardSlides.getSlides()[2] // Uplift slide
  
  Logger.log("Duplicate uplift slide to match # of exposed");
  for (var i = 1; i < nExposed; i++){
    slide.duplicate()
  }

  // Create slide for each activity

  for (var i = 0; i < nExposed; i++) {
    // Insert uplifts to each expose slide
    var slide = scorecardSlides.getSlides()[2 + i]
    var slidebyID = scorecardSlides.getSlideById(slide.getObjectId())

    // Set scorecard sheet
    var sheet = dataSpreadsheet.getSheetByName("Scorecard");
    var sourceSheet = dataSpreadsheet.getSheetByName("Source>>");

    // Get number of control respondents
    var nResponsesCtrl = sourceSheet.getRange("J3").getValue();

    // Get all values; in same order as spreadsheet
    var activityID = sheet.getRange(4 + i, 1, 1, 1).getValue();
    var activityName = sheet.getRange(4 + i, 2, 1, 1).getValue().toUpperCase();
    var mainBrand = sheet.getRange(4 + i, 3, 1, 1).getValue().toUpperCase();
    var nResponsesExp = sheet.getRange(4 + i, 4, 1, 1).getValue() + nResponsesCtrl;
    

    var avgImpressionS = sheet.getRange(4 + i, 5, 1, 1).getValue() + "s";
    var aidedAwareness = (sheet.getRange(4 + i, 6, 1, 1).getValue() * 100).toFixed(2) + "%";
    var fastYesUplift = (sheet.getRange(4 + i, 7, 1, 1).getValue() * 100).toFixed(2) + "%";
    var favourabilityUplift = (sheet.getRange(4 + i, 8, 1, 1).getValue() * 100).toFixed(2) + "%";
    var considerUplift = (sheet.getRange(4 + i, 9, 1, 1).getValue() * 100).toFixed(2) + "%";
    var purchaseUplift = (sheet.getRange(4 + i, 10, 1, 1).getValue() * 100).toFixed(2) + "%";

    var customUplift = (sheet.getRange(4 + i, 34, 1, 1).getValue() * 100).toFixed(2) + "%";

    var expressionAvg = (sheet.getRange(4 + i, 38, 1, 1).getValue() * 100).toFixed(2) + "%";
    var expressionMax = (sheet.getRange(4 + i, 39, 1, 1).getValue() * 100).toFixed(2) + "%";

    var customTrait = sheet.getRange(4 + i, 33, 1, 1).getValue().toUpperCase(); // Get custom trait value

    var bNResponsesExp = dataSpreadsheet.getSheetByName("Scorecard Benchmark").getRange("B2").getValue();

    // Heading & Footer
    Logger.log("Set Heading and Footer")
    slidebyID.replaceAllText('{{handle}}', activityName);
    slidebyID.replaceAllText('{{mainBrand}}', mainBrand);
    slidebyID.replaceAllText('{{activityID}}', activityID);
    slidebyID.replaceAllText('{{date}}', date);
    slidebyID.replaceAllText('{{nResponsesExp}}', nResponsesExp);
    slidebyID.replaceAllText('{{bNResponsesExp}}', bNResponsesExp);
    slidebyID.replaceAllText('{{customTrait}}', customTrait);

    // Funnel Uplifts
    Logger.log("Set Funnel Uplifts")
    slidebyID.replaceAllText('{{avgImpressionS}}', avgImpressionS);
    slidebyID.replaceAllText('{{aidedAwareness}}', aidedAwareness);

    slidebyID.replaceAllText('{{expressionAvg}}', expressionAvg);
    slidebyID.replaceAllText('{{expressionMax}}', expressionMax);
    slidebyID.replaceAllText('{{fastYesUplift}}', fastYesUplift);
    slidebyID.replaceAllText('{{customUplift}}', customUplift);
    slidebyID.replaceAllText('{{favUplift}}', favourabilityUplift);
    slidebyID.replaceAllText('{{considerUplift}}', considerUplift);
    slidebyID.replaceAllText('{{purchaseUplift}}', purchaseUplift);

  // Insert same benchmarks
    var benchmark = dataSpreadsheet.getRange("Project!G11").getValue();
    var nActivitiesBenchmark = dataSpreadsheet.getRange("Project!H11").getValue();

    slidebyID.replaceAllText('{{benchmark}}', benchmark);
    slidebyID.replaceAllText('{{nActivitiesBenchmark}}', nActivitiesBenchmark);
    
    Logger.log("Get benchmark diff value")
    var bAvgImpressionS = sheet.getRange(4 + i, 17, 1, 1).getValue();
    var bAidedAwareness = sheet.getRange(4 + i, 18, 1, 1).getValue();
    var bFastYesUplift = sheet.getRange(4 + i, 19, 1, 1).getValue();
    var bFavourabilityUplift = sheet.getRange(4 + i, 20, 1, 1).getValue();
    var bConsiderUplift = sheet.getRange(4 + i, 21, 1, 1).getValue();
    var bPurchaseUplift = sheet.getRange(4 + i, 22, 1, 1).getValue();

    var bAvgImpressionSColor = sheet.getRange(4 + i, 23, 1, 1).getValue();
    var bAidedAwarenessColor = sheet.getRange(4 + i, 24, 1, 1).getValue();
    var bFastYesUpliftColor = sheet.getRange(4 + i, 25, 1, 1).getValue();
    var bFavourabilityUpliftColor = sheet.getRange(4 + i, 26, 1, 1).getValue();
    var bConsiderUpliftColor = sheet.getRange(4 + i, 27, 1, 1).getValue();
    var bPurchaseUpliftColor = sheet.getRange(4 + i, 28, 1, 1).getValue();

    slidebyID.replaceAllText('{{bAvgImpressionS}}', bAvgImpressionS);
    slidebyID.replaceAllText('{{bAidedAwareness}}', bAidedAwareness);
    slidebyID.replaceAllText('{{bFastYesUplift}}', bFastYesUplift);
    slidebyID.replaceAllText('{{bFavourabilityUplift}}', bFavourabilityUplift);
    slidebyID.replaceAllText('{{bConsiderUplift}}', bConsiderUplift);
    slidebyID.replaceAllText('{{bPurchaseUplift}}', bPurchaseUplift);
    
    Logger.log("Set benchmark difference color")
    var slideTable = slide.getTables()[0]
    // cellBAvgImpressionS
    slideTable.getCell(1,2).getText().getTextStyle().setForegroundColor(bAvgImpressionSColor);
    // cellBAidedAwareness
    slideTable.getCell(2,2).getText().getTextStyle().setForegroundColor(bAidedAwarenessColor);
    // cellBFastYesUplift
    slideTable.getCell(4,2).getText().getTextStyle().setForegroundColor(bFastYesUpliftColor);
    // cellBFavourabilityUplift
    slideTable.getCell(6,2).getText().getTextStyle().setForegroundColor(bFavourabilityUpliftColor);
    // cellBConsiderUplift
    slideTable.getCell(7,2).getText().getTextStyle().setForegroundColor(bConsiderUpliftColor);
    // cellBPurchaseUplift
    slideTable.getCell(8,2).getText().getTextStyle().setForegroundColor(bPurchaseUpliftColor);

    // Insert gif
    try{
    var gifFile = getStimGif(activityID, scorecardFolder.getId())
    slide.insertImage(gifFile).setLeft(464.009).setTop(63.594).setHeight(296.007).setWidth(148.511).sendToBack();

    } catch(error){
      Logger.log("Unable to insert gif for " + activityName + ", please insert manually. Report will continue to generate.")
      sendSlack("Unable to insert gif for " + activityName + ", please insert manually. Report will continue to generate.")
    }
  }
}

function getPublishedScorecards() {
  Logger.log("Start getPublishedScorecards");

  const folder = DriveApp.getFolderById("1L-39een3hF0NmHL_ddZbHMR_5FAjPNv5");
  const subfolders = folder.getFolders();


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Published Scorecards");
  sheet.clear();

  const header = ["Scorecard Name", "Date Created", "Folder Url"]
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");

  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const name = subfolder.getName();
    const date = subfolder.getDateCreated();
    const url = subfolder.getUrl();
    const lastRow = sheet.getLastRow();

    sheet.getRange(lastRow + 1, 1).setValue(name);

    sheet.getRange(lastRow + 1, 2).setValue(date);

    sheet.getRange(lastRow + 1, 3).setValue(url);
  }

  Logger.log("End getPublishedScorecards");
}
