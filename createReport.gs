function createReport() {
  try {
    logMsg("START create report")
    // Import data source to template slide
    importDataSource();

    calculateBenchmarkDiff();

    // Create report folder and template copies
    var reportFiles = createReportFolder();
    var dataSpreadsheet = reportFiles[0]
    var scorecardSlides = reportFiles[1]
    var scorecardFolder = reportFiles[2]
    var scorecardName = reportFiles[3]

    Logger.log(dataSpreadsheet);
    Logger.log(scorecardSlides);
    Logger.log(scorecardFolder);


    // Create activity slides
    createActivitySlides(dataSpreadsheet, scorecardSlides, scorecardFolder);

    // Get all published scorecards
    getPublishedScorecards();

    logMsg("JOB DONE: Scorecard generated")
    sendSlack(scorecardName + "   " + scorecardFolder.getUrl())


  } catch (error) {
    logMsg("Error generating scorecard: " + error + "Incomplete results. Please regenerate.")
    sendSlack("Error generating scorecard: " + error + "Incomplete results. Please regenerate.")
  }


}

function copyReport() {
  try {
    logMsg("START copy folder")
    // Get scorecard folder URL
    logMsg("Get from folder");
    var fromFolderID = SpreadsheetApp.getActive().getRange("Dashboard!B15").getValue();
    var fromFolder = DriveApp.getFolderById(fromFolderID);
    Logger.log("from folder ID: " + fromFolderID)
    Logger.log("from folder: " + fromFolder);

    // Destination folder ID
    logMsg("Create copy folder to destination");
    var toFolderID = "1l6nl01lydrK-cytFUnN86iKmBXvzQjva"
    var clientFolder = DriveApp.getFolderById(toFolderID);
    var toFolder = clientFolder.createFolder(fromFolder);
    Logger.log("to folder: " + toFolder);

    // Copy folder content files
    logMsg("Copy folder content files to destination folder")
    var files = fromFolder.getFiles();
    while (files.hasNext()) {
      Logger.log("copy file in folder");
      var file = files.next();
      var newFile = file.makeCopy(toFolder)
      newFile.setName(file.getName())
    }
    
    logMsg("END copy folder")
    sendSlack("Copy of " + toFolder + ":  " + toFolder.getUrl());
    
  } catch (error) {
    Logger.log(error)
    sendSlack("Copy folder error " + toFolder + ":  " + error);
  }
}
