
var SEARCH_QUERY = "to:me subject:(Kululasku lähetetty kopiona sinulle)";
var RECEIPT_SHEET = "Kuitit kululasku.fi:stä";
var _ = LodashGS.load();

// Credit: https://gist.github.com/oshliaer/70e04a67f1f5fd96a708

function getEmails_(q) {

  // Get conent of receipt emails
  var emails = [];
  var threads = GmailApp.search(q);
  for (var i in threads) {
    var msgs = threads[i].getMessages();
    for (var j in msgs) {
      emails.push([msgs[j].getBody().replace(/<.*?>/g, '\n')
                   .replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')
                  ]);
    }
  }
  
  return emails;
}

// Parse emails from an nx1 two-dimensional array of multiline strings to an 2d array with columns the data about to be imported
function parseExpenses(emails) {
  
  var expenses = [];
  var row_index = 0;

  for (var i in emails) {
    
    // Every element of emails contains an array with length one in whihc the message is
    var msg = emails[i][0].split(/\r/);
    var nro_of_expenses = (msg.length - 13)/2;

    for (var j=0; j < nro_of_expenses; j++) {
      expenses[row_index] = [];
      // Date
      expenses[row_index][0] = msg[8+2*j].split(/\s/)[0];
      // Explanation
      expenses[row_index][1] = msg[5];
      // Amount
      expenses[row_index][2] = msg[8+2*j].match(/\w+,\w\w$/);
      // Expense spesific explanation
      expenses[row_index][3] = msg[9+2*j];
      // Name of the person issued an expense
      expenses[row_index][4] = msg[2].split(/Nimi:\s+/)[1];
      row_index += 1;
    }
  }
  return expenses;
}

// Return an array of hashes.
// Each hash is a concatenation of all the items in one row
function parseToHashes (array2d) {
  hashes = new Array(array2d.length);
  for (var i = 0; i < hashes.length; i++){
    hashes[i] = "";
    for (var j = 0; j < array2d[0].length; j++) {
      hashes[i] += array2d[i][j]
    }
  }
  return hashes;
}


function appendData(sheet, array2d, overwrite) {
  if (array2d.length == 0) {
    return
  }
  
  if (overwrite) {
    // Clear old data
    sheet.getRange(2,1,sheet.getLastRow()-1, array2d[0].length).clear();
    // Append new data
    sheet.getRange(2, 1, array2d.length, array2d[0].length).setValues(array2d);
  }
  else {
    sheet.getRange(sheet.getLastRow() + 1, 1, array2d.length, array2d[0].length).setValues(array2d);
  }
}


function saveEmails() {

  Logger.log("start");
  var array2d = getEmails_(SEARCH_QUERY)
  
  // If any emails was found
  if (array2d) {
    Logger.log(array2d);
    var expenses = parseExpenses(array2d);
    var receiptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIPT_SHEET);
    Logger.log(expenses.length);
    
    // Get hashes of current receipts from the expenses tab
    var nOfReceipts = receiptSheet.getLastRow()-1;
    var existingReceiptHashes = [] 
    if (nOfReceipts <= 0) {
      existingReceiptHashes = [];
    } else {
      existingReceiptHashes = receiptSheet.getRange(2, 1, nOfReceipts, 1).getValues();
    }
    
    
    // Identify new receips that are not yet imported to the sheet
    var fetchedReceiptsHashes = parseToHashes(expenses);
    
    // Prepare existing hashes for comparision by parsing them from 2d array to 1d array
    var receiptHashes = parseToHashes(existingReceiptHashes);
    var newExpenses = []
    var newHashNumbers = []
    
    for (i=0; i < fetchedReceiptsHashes.length; i++) {
      if (_.indexOf(receiptHashes, fetchedReceiptsHashes[i]) == -1) {
        newExpenses[newExpenses.length] = expenses[i];
        newHashNumbers[newHashNumbers.length] = i;
      }
    }
    
    
    // Modify expenses from emails: replace commas with dots for number
    expenses = expenses.map(function (row) {
      row[2] = row[2].toString().replace(",",".");
      return row;
    });
    
    
    // Add indexes for new data to be appended
    for (i = 0; i < newExpenses.length; i++){
      newExpenses[i] = _.concat(fetchedReceiptsHashes[newHashNumbers[i]], nOfReceipts+i+1, newExpenses[i]);
    }
    
    
    // Append new expenses to the spreadsheet
    appendData(receiptSheet, newExpenses);
    Logger.log("Added " + newExpenses.length + " expenses with " + (expenses[0].length+2) + " columns each.");
  }
}
