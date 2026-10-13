// Global Variables
const API_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOiIyM2I3ZmU4MC0xNjhlLTExZWItYmQwMi1jNzc0NDg1ZDcxNzciLCJzdWIiOiJTSE9VVE9VVF9BUElfVVNFUiIsImlhdCI6MTYwMzYwODUzMCwiZXhwIjoxOTE5MTQxMzMwLCJzY29wZXMiOnsiYWN0aXZpdGllcyI6WyJyZWFkIiwid3JpdGUiXSwibWVzc2FnZXMiOlsicmVhZCIsIndyaXRlIl0sImNvbnRhY3RzIjpbInJlYWQiLCJ3cml0ZSJdfSwic29fdXNlcl9pZCI6IjMxNzMiLCJzb191c2VyX3JvbGUiOiJ1c2VyIiwic29fcHJvZmlsZSI6ImFsbCIsInNvX3VzZXJfbmFtZSI6IiIsInNvX2FwaWtleSI6Im5vbmUifQ.hW02s9uZ_nOPYVhrPtmMYcp4wB33IcIHDlUmOuopmrA';

const SOURCE_NAME = 'BISEKA';

function scriptRun() {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Send");
  var lRow = currentSheet.getLastRow();
  
  if (lRow <= 1) { return; } // Skip if only headers exist

  var data = currentSheet.getRange(2, 1, lRow - 1, 4).getValues();
  currentSheet.getRange(2, 1, lRow - 1, 4).clearContent();

  // FIX: Filter out all the empty, invisible, or formatted rows
  var validData = data.filter(row => row[0].toString().trim() !== "" && row[2].toString().trim() !== "");
  if (validData.length === 0) return;

  var requests = [];
  var date = new Date();

  validData.forEach(row => {
    var order = row[0];
    var agent = row[1];
    var contact = '94' + row[2];
    var info = row[3];
    
    var lastLine = (agent == 'Safeway') ? '1-3 working days\n\n Rider\' contact: ' + info : '2-7 working days\n\n Check status: ' + info;
    
    var payload = {
      'source': SOURCE_NAME,
      'destinations': [contact],
      'content': { 'sms': 'Dear Customer, your order # ' + order + ' dispatched & arrive within ' + lastLine },
      'transports': ['sms'],
    };

    requests.push({
      'url': 'https://api.getshoutout.com/coreservice/messages',
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': 'Apikey ' + API_KEY,
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    });
  });

  // Send all valid requests simultaneously
  var responses = UrlFetchApp.fetchAll(requests);
  var newArray = [];

  validData.forEach((row, index) => {
    var reqData;
    try {
      reqData = JSON.parse(responses[index].getContentText());
    } catch (e) {
      reqData = { error: 'API parse error', details: e.toString() };
    }
    row.push(date);
    row.push(JSON.stringify(reqData)); // Write to Sheets safely as a string
    newArray.push(row);
  });

  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  logSheet.getRange(logSheet.getLastRow() + 1, 1, newArray.length, 6).setValues(newArray);
}

function failedOrderReq() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Failed Order');
  var lRow = ws.getLastRow();
  
  if (lRow <= 1) { return; }

  var data = ws.getRange(2, 1, lRow - 1, 2).getValues();
  ws.getRange(2, 1, lRow - 1, 5).clearContent();

  // FIX: Filter out blank rows
  var validData = data.filter(row => row[0].toString().trim() !== "" && row[1].toString().trim() !== "");
  if (validData.length === 0) return;

  var requests = [];
  var date = new Date();

  validData.forEach(row => {
    var contact = '94' + row[0];
    var orderno = row[1];
    var payload = {
      'source': SOURCE_NAME,
      'destinations': [contact],
      'content': { 'sms': 'Dear Valued Customer,\nThe order # ' + orderno + ' you submitted through our Website was not successful.\n\nPlace your order now on WhatsApp 0710728124' },
      'transports': ['sms'],
    };

    requests.push({
      'url': 'https://api.getshoutout.com/coreservice/messages',
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': 'Apikey ' + API_KEY,
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    });
  });

  var responses = UrlFetchApp.fetchAll(requests);
  var newData = [];

  validData.forEach((row, index) => {
    var res;
    try {
      res = JSON.parse(responses[index].getContentText());
    } catch (e) {
      res = { status: 'failed' };
    }
    row.push(date);
    row.push(JSON.stringify(res));
    row.push(res.status || 'failed');
    newData.push(row);
  });

  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Failed Order - Log");
  logSheet.getRange(logSheet.getLastRow() + 1, 1, newData.length, 5).setValues(newData);

  var error = [];
  newData.forEach(row => {
    if (row[4] != '1001') error.push([row[0], row[1], 'error']);
  });
  
  if (error.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, error.length, 3).setValues(error);
  }
}

function outOfStock() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Out of Stock');
  var lRow = ws.getLastRow();
  
  if (lRow <= 1) { return; }

  var noOfCols = 3;
  var data = ws.getRange(2, 1, lRow - 1, noOfCols).getValues();
  ws.getRange(2, 1, lRow - 1, noOfCols + 1).clearContent();

  // FIX: Filter out blank rows
  var validData = data.filter(row => row[0].toString().trim() !== "" && row[1].toString().trim() !== "");
  if (validData.length === 0) return;

  var requests = [];
  var date = new Date();

  validData.forEach(row => {
    var contact = '94' + row[0];
    var orderno = row[1];
    var newdeldate = row[2];
    var msg = `Dear Customer, Your order # ${orderno}, have not yet dispatched due to running out of stock.\n\nWe will dispatch your order on ${newdeldate}.`;

    var payload = {
      'source': SOURCE_NAME,
      'destinations': [contact],
      'content': { 'sms': msg },
      'transports': ['sms'],
    };

    requests.push({
      'url': 'https://api.getshoutout.com/coreservice/messages',
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': 'Apikey ' + API_KEY,
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    });
  });

  var responses = UrlFetchApp.fetchAll(requests);
  var newData = [];

  validData.forEach((row, index) => {
    var res;
    try {
      res = JSON.parse(responses[index].getContentText());
    } catch (e) {
      res = { status: 'failed' };
    }
    
    row.push(date);
    row.push(JSON.stringify(res)); 
    row.push(res.status || 'failed');
    newData.push(row);
  });

  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Out of Stock - Log");
  logSheet.getRange(logSheet.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);

  var resposeStatus = newData[0].length - 1;
  var error = [];
  newData.forEach(row => {
    if (row[resposeStatus] != '1001') error.push([row[0], row[1], row[2], 'error']);
  });

  if (error.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, error.length, error[0].length).setValues(error);
  }
}
