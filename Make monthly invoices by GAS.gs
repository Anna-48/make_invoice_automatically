/** 
  This is the link of Invoice Template: https://docs.google.com/spreadsheets/d/1G84U-NU6oaEWt4p3ky61MZFbkGY6ZtRnfbYMK-PuCaM/edit?usp=sharing
*/

// ADD A BUTTON IN THE TASKBAR
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invoice')
                .addItem('MERCHANTS WITH TRANSACTION + MONTHLY FEE', 'merchant_with_transaction_and_monthly_fee') 
                .addItem('MERCHANTS WITH JUST MONTHLY FEE', 'merchant_with_just_monthly_fee')
                .addToUi(); 
};

// INVOICE TO MERCHANTS WITH TRANSACTION + MONTHLY FEE
function merchant_with_transaction_and_monthly_fee(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invTempSheet = ss.getSheetByName("INV TEMPLATE");
  var invListSheet = ss.getSheetByName("INV list");
  var lastMerchant = invListSheet.getRange('A2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  var merchants = invListSheet.getRange(`A3:B${lastMerchant}`).getValues();

  for(var i = 0; i < merchants.length; i++){
    // SET MERCHANT
    invTempSheet.activate();
    invTempSheet.getRange('B13').setValue(merchants[i][1]);

    // DUPLICATE SHEET
    var fileName = merchants[i][0];
    var newSheet = ss.duplicateActiveSheet();
    newSheet.setName(fileName);
    
    // CONVERT FORMULAS (RELATED TO DATE/TIME) TO VALUES
    var range1 = newSheet.getRange('B9');
    range1.copyTo(range1, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    var range2 = newSheet.getRange('F12');
    range2.copyTo(range2, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range3 = newSheet.getRange('F15');
    range3.copyTo(range3, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range4 = newSheet.getRange('B19:D20');
    range4.copyTo(range4, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    var range5 = newSheet.getRange('B30');
    range5.copyTo(range5, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range6 = newSheet.getRange('B32:G42');
    range6.copyTo(range6, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // FIND ROWs
    var maxRow = newSheet.getMaxRows();
    var totalCell = newSheet.getRange(`C${maxRow}`).getNextDataCell(SpreadsheetApp.Direction.UP);
    var totalRow = totalCell.getRow();
    var lastRow = totalCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    // DELETE BLANKS
    var startRow = lastRow + 2;
    var numRows = totalRow - startRow;
    if (startRow < totalRow - 1) {
      newSheet.deleteRows(startRow,numRows);
    };
  };
}

// INVOICE TO MERCHANTS WITH JUST MONTHLY FEE
function merchant_with_just_monthly_fee(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invTempSheet = ss.getSheetByName("INV TEMPLATE");
  var invListSheet = ss.getSheetByName("INV list");
  var lastMerchant = invListSheet.getRange('D2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  var merchants = invListSheet.getRange(`D3:E${lastMerchant}`).getValues();

  for(var i = 0; i < merchants.length; i++){
    // SET MERCHANT
    invTempSheet.activate();
    invTempSheet.getRange('B13').setValue(merchants[i][1]);

    // DUPLICATE SHEET
    var fileName = merchants[i][0];
    var newSheet = ss.duplicateActiveSheet();
    newSheet.setName(fileName);
    
    // CONVERT FORMULAS (RELATED TO DATE/TIME) TO VALUES
    var range1 = newSheet.getRange('B9');
    range1.copyTo(range1, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    var range2 = newSheet.getRange('F12');
    range2.copyTo(range2, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range3 = newSheet.getRange('F15');
    range3.copyTo(range3, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range4 = newSheet.getRange('B19:D20');
    range4.copyTo(range4, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    var range5 = newSheet.getRange('B30');
    range5.copyTo(range5, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range6 = newSheet.getRange('B32:G42');
    range6.copyTo(range6, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // FIND ROWs
    var maxRow = newSheet.getMaxRows();

    // DELETE BLANKS
    var startRow = 30;
    var numRows = maxRow - startRow;
    newSheet.deleteRows(startRow,numRows); 
  };
}
