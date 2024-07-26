function onOpen(e) {
  // Add a custom menu to the spreadsheet.
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Scorecard')
      .addItem('Create Scorecard', 'createReport')
      .addItem('Copy Scorecard Folder', 'copyReport')
      .addToUi();
}

function logMsg(message){
  const range = SpreadsheetApp.getActive().getRange("Dashboard!B19");
  const msg = Logger.log(message).getLog();
  var logMsg = range.setValue(msg);
  return logMsg
}


function sendSlack(message) {
  // Send slack notification
  Logger.log("Send slack notification");
  const url = "dummyurl
  const params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      "text": message
    })
  }
  const sendMsg = UrlFetchApp.fetch(url, params)
  var respCode = sendMsg.getResponseCode()
}
