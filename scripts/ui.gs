function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Manage Bookings')
      .addItem('Add Bookings', 'showSidebar')
      .addItem('Cancel Bookings', 'cancel')
      .addToUi();
}

// User Interfaces

function doGet(e) {
  return HtmlService.createTemplateFromFile('requestForm')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function getScriptURL() {
  let url = ScriptApp.getService().getUrl();
  return url;
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar')
      .evaluate()
      .setTitle('Enter booking')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function addSidebarTab(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// Dashboard operations

function view() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = sheet.getActiveCell();
  var reqId = cell.getValue();
  var row = sheet.getRange('M' + cell.getRow()).getValue();
  var requests = sheet.getSheetByName('Requests');
  var header = requests.getRange(2,4,1,requests.getMaxColumns()).getValues()[0];
  var request = requests.getRange(row, 4, 1, requests.getMaxColumns()).getDisplayValues()[0];
  var mask = request.map(item => item !== "");
  request = request.filter((item, i) => mask[i]);
  header = header.filter((item, i) => mask[i]);
  var summary = sheet.getSheetByName('Dashboard');
  summary.getRange('B4:C40').clearContent();
  summary.getRange('C3').setValue(reqId);
  for(var i = 0; i < request.length; i++){
    summary.getRange(4+i,2,1,1).setValue(header[i]);
    summary.getRange(4+i,3,1,1).setValue(request[i]);
  };
}

function clear() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('B4:C40').clearContent();
  sheet.getRange('C3').clearContent();
};

// Booking operations

function start() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = sheet.getActiveCell();
  var reqId = cell.getValue();
  var summary = sheet.getSheetByName('Dashboard')
  var row = summary.getRange('M' + cell.getRow()).getValue();
  var requests = sheet.getSheetByName('Requests');
  if (reqId == requests.getRange(row,3,1,1).getValue()) {
    requests.getRange(row,1,1,1).setValue('On going');
  }
};

function finish() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = sheet.getActiveCell();
  var reqId = cell.getValue();
  var summary = sheet.getSheetByName('Dashboard')
  var row = summary.getRange('M' + cell.getRow()).getValue();
  var requests = sheet.getSheetByName('Requests');
  if (reqId == requests.getRange(row,3,1,1).getValue()) {
    requests.getRange(row,1,1,1).setValue('Booked');
  }
};

function cancel() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var requests = sheet.getSheetByName('Requests')
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var status = requests.getRange('A'+row).getValue();
  requests.getRange('A' + row).setValue('Cancelled');

  var reqId = requests.getRange('C'+row).getValue();
  var type = requests.getRange('D'+row).getValue().split("+");
  Logger.log(type)

  if (status == 'Booked') {
    for (i in type) {
      var logBook = sheet.getSheetByName(type[i].trim());
      var lastRow = logBook.getLastRow();
      while(lastRow){
        if(reqId==(logBook.getRange('A' + lastRow).getValue())){
          logBook.getRange('B' + lastRow).setValue('Cancelled');
          break;
        };
      lastRow--;
      };
    }
  }

};

