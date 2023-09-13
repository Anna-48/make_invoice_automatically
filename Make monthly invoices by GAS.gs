/** @OnlyCurrentDoc */

function makeInvoice(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invTempSh = ss.getSheetByName("INV TEMPLATE");
  var sumSh = ss.getSheetByName("SUM");
  var LastRow = sumSh.getRange('L2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  var merchants = sumSh.getRange(2,12,LastRow-1,2).getValues();

  for(var i = 0; i < merchants.length; i++){
    // SET MERCHANT
    invTempSh.activate();
    invTempSh.getRange('A12').setValue(merchants[i][1]);

    // DUPLICATE SHEET
    var fileName = merchants[i][0];
    var newSh = ss.duplicateActiveSheet();
    newSh.setName(fileName);
    
    // CONVERT FORMULAS TO VALUES
    var range1 = newSh.getRange('H7:H8');
    range1.copyTo(range1, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    var range2 = newSh.getRange('C16:E17');
    range2.copyTo(range2, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range3 = newSh.getRange('A44');
    range3.copyTo(range3, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    var range4 = newSh.getRange('A47:H77');
    range4.copyTo(range4, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    // DELETE BLANKS
    var startRow = newSh.getRange("B80").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 2;
    var numRows = 79 - startRow;
    if (startRow < 79) {
      newSh.deleteRows(startRow,numRows)
    };
  };
}
