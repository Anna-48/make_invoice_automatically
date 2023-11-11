/** @OnlyCurrentDoc */

function onOpen() {
 var ui = SpreadsheetApp.getUi().createMenu('Custom')
                        .addItem('Process data', 'processData') 
                        .addItem('1st of month', 'firstOfMonth')
                        .addToUi(); 
}


// xử lý dữ liệu cho các ngày trong tháng
function processData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getActiveSheet();

// FIND THE FIRST ROW HAVING TRANSACTION
  var textFinder = dataSheet.createTextFinder('ACCOUNT NBR').findNext().getRow();
  var firstRow = textFinder + 1;

// CHANGE NAME OF SHEET
  var reportDate = dataSheet.getRange('B4').getDisplayValue();
  var day = reportDate.toString().slice(0,2);
  var month = reportDate.toString().slice(3,5);
  var year = reportDate.toString().slice(6);
  var termID = dataSheet.getRange(`G${firstRow}`).getValue();
  var merId = termID.toString().slice(5);
  dataSheet.setName(`${day}${month}-${merId}`);

// CHANGE FORMAT OF SHEET
  dataSheet.getRange(1, 1, dataSheet.getMaxRows(), dataSheet.getMaxColumns())
           .setFontFamily('Calibri').setFontSize(11);

  dataSheet.setColumnWidth(4, 26);
  dataSheet.setColumnWidth(6, 26);
  dataSheet.setColumnWidth(8, 26);
  dataSheet.setColumnWidth(10, 26);
  dataSheet.setColumnWidth(12, 26);
  dataSheet.setColumnWidth(14, 26);

// DELETE BLANK ROWS
  dataSheet.deleteRows(100,900)

// COPY DATA
  dataSheet.getRange(`A${firstRow}:O${firstRow}`).activate();
  var data = spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  
// FIND THE SUM SHEET 
  var dataShID = dataSheet.getIndex();
  var sumSheet = spreadsheet.getSheets()[dataShID - 2];

// REMOVE PROTECTION
  sumSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove();

// PASTE DATA
  var lastRow1 = sumSheet.getLastRow();
  var start = lastRow1 + 1;
  data.copyTo(sumSheet.getRange(start,1),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

// CLEAR THE TOTAL ROW
  var lastRow2 = sumSheet.getLastRow();
  var end = lastRow2 - 1;
  sumSheet.getRange(lastRow2,1,1,15).clear();

// FILL FORMULAS
  var lastCol = sumSheet.getLastColumn();
  var numofCol = lastCol - 15;
  sumSheet.getRange(lastRow1 - 1,16,1,numofCol).copyTo(sumSheet.getRange(start,16),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sumSheet.getRange(start,16,1,numofCol).autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

// SET FORMULAS
  sumSheet.getRange(lastRow2,3).setFormula(`=SUM(C${start}:C${end})`);
  sumSheet.getRange(lastRow2,18).setFormula(`=SUM(R${start}:R${end})`);
  sumSheet.getRange(lastRow2,23).setFormula(`=SUM(W${start}:W${end})`);

// CHANGE FORMAT
  sumSheet.getRange(`${lastRow1}:${lastRow1}`).copyTo(sumSheet.getRange(`${lastRow2}:${lastRow2}`), SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);

  sumSheet.getRangeList([`C${lastRow2}`,`R${lastRow2}`,`W${lastRow2}`]).activate();

// special case: 34000313
  if (termID == 34000313) {
      var rrCell = sumSheet.getRange(`Y${lastRow2}`);
      sumSheet.getRange(`W${lastRow2}`).copyTo(rrCell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      rrCell.activate();
  }

// SET PROTECTION
  sumSheet.protect().setWarningOnly(true);

// // COPY TO THE REPORT
//   // OPEN SHEET
//   var reportFile = SpreadsheetApp.openById('1nNfL6ZHIVzGkb-AwTa8JvjMxhZ6LKcveiF7Y7dMnq1A');
//   var date = new Date(`${month}/${day}/${year}`);
//   var thisMonth = date.toLocaleString('en-US', {month: 'short', year: '2-digit'});
//   var reportSheet = reportFile.getSheetByName(thisMonth);

//   // GET LAST ROW
//   var maxRow = reportSheet.getMaxRows();
//   var endRow = reportSheet.getRange(`B${maxRow}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
//   var fillRow = endRow + 1;

//   // COPY DATA
//   var trxAmt = sumSheet.getRange(`C${lastRow2}`).getValue();
//   reportSheet.getRange(`E${fillRow}`).setValue(trxAmt);
//   var bankAmt = sumSheet.getRange(`R${lastRow2}`).getValue();
//   reportSheet.getRange(`F${fillRow}`).setValue(bankAmt);
//   var merAmt = sumSheet.getRange(`W${lastRow2}`).getValue();
//   reportSheet.getRange(`G${fillRow}`).setValue(merAmt);
//   reportSheet.getRange(`B${fillRow}`).setValue(termID);
};

// xử lý dữ liệu cho ngày đầu tiên có báo cáo
function firstOfMonth(){
  var spreadsheet = SpreadsheetApp.getActive();
  var dataSheet = spreadsheet.getActiveSheet();

// FIND THE FIRST ROW HAVING TRANSACTION
  var textFinder = dataSheet.createTextFinder('ACCOUNT NBR').findNext().getRow();
  var firstRow = textFinder + 1;

// CHANGE NAME OF SHEET
  var reportDate = dataSheet.getRange('B4').getDisplayValue();
  var day = reportDate.toString().slice(0,2);
  var month = reportDate.toString().slice(3,5);
  var year = reportDate.toString().slice(6);
  var termID = dataSheet.getRange(`G${firstRow}`).getValue();
  var merId = termID.toString().slice(5);
  dataSheet.setName(`${day}${month}-${merId}`);

// CHANGE FORMAT OF SHEET
  dataSheet.getRange(1, 1, dataSheet.getMaxRows(), dataSheet.getMaxColumns())
           .setFontFamily('Calibri').setFontSize(11);

  dataSheet.setColumnWidth(4, 26);
  dataSheet.setColumnWidth(6, 26);
  dataSheet.setColumnWidth(8, 26);
  dataSheet.setColumnWidth(10, 26);
  dataSheet.setColumnWidth(12, 26);
  dataSheet.setColumnWidth(14, 26);

// DELETE BLANK ROWS
  dataSheet.deleteRows(100,900)
  
// COPY DATA
  dataSheet.getRange(`A${firstRow}:O${firstRow}`).activate();
  var data = spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

// FIND THE SUM SHEET  
  var dataShID = dataSheet.getIndex();
  var sumSheet = spreadsheet.getSheets()[dataShID - 2];

// REMOVE PROTECTION
  sumSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove();

// PASTE DATA
  var lastRow1 = sumSheet.getLastRow();
  data.copyTo(sumSheet.getRange(lastRow1,1),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

// CLEAR THE TOTAL ROW
  var lastRow2 = sumSheet.getLastRow();
  var end = lastRow2 - 1;
  sumSheet.getRange(lastRow2,1,1,15).clear();

// FILL FORMULAS
  var lastCol = sumSheet.getLastColumn();
  var numofCol = lastCol - 15;
  sumSheet.getRange(lastRow1,16,1,numofCol).autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

// SET FORMULAS
  sumSheet.getRange(lastRow2,3).setFormula(`=SUM(C${lastRow1}:C${end})`).setBackground('#d9ead3');
  sumSheet.getRange(lastRow2,18).setFormula(`=SUM(R${lastRow1}:R${end})`).setBackground('#ffff00');
  sumSheet.getRange(lastRow2,23).setFormula(`=SUM(W${lastRow1}:W${end})`).setBackground('#deeaf6');

// CHANGE FORMAT
  sumSheet.getRangeList([`C${lastRow2}`,`R${lastRow2}`,`W${lastRow2}`])
          .setFontFamily('Calibri').setFontSize(11)
          .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
          .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DOUBLE)
          .setFontWeight('bold')
          .setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)')
          .activate();

// special case: 34000313
  if (termID == 34000313) {
      var rrCell = sumSheet.getRange(`Y${lastRow2}`);
      sumSheet.getRange(`W${lastRow2}`).copyTo(rrCell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      rrCell.activate();
  }

// SET PROTECTION
  sumSheet.protect().setWarningOnly(true);

// // COPY TO THE REPORT
//   // OPEN SHEET
//   var reportFile = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1nNfL6ZHIVzGkb-AwTa8JvjMxhZ6LKcveiF7Y7dMnq1A/edit#gid=2065604187');
//   var date = new Date(`${month}/${day}/${year}`);
//   var thisMonth = date.toLocaleString('en-US', {month: 'short', year: '2-digit'});
//   var reportSheet = reportFile.getSheetByName(thisMonth);

//   // GET LAST ROW
//   var maxRow = reportSheet.getMaxRows();
//   var endRow = reportSheet.getRange(`B${maxRow}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
//   var fillRow = endRow + 1;

//   // COPY DATA
//   var trxAmt = sumSheet.getRange(`C${lastRow2}`).getValue();
//   reportSheet.getRange(`E${fillRow}`).setValue(trxAmt);
//   var bankAmt = sumSheet.getRange(`R${lastRow2}`).getValue();
//   reportSheet.getRange(`F${fillRow}`).setValue(bankAmt);
//   var merAmt = sumSheet.getRange(`W${lastRow2}`).getValue();
//   reportSheet.getRange(`G${fillRow}`).setValue(merAmt);
//   reportSheet.getRange(`B${fillRow}`).setValue(termID);

};


/*
tổng hợp báo cáo cả tháng của merchants
các bước cần làm trước khi dùng lệnh: 
1. mở sheet 'Sum of month'
2. lấy số lượng merchants có transaction trong tháng bằng cách dùng pivot table bên report file 
3. copy list of merchants vào ô Y2
4. tại sheet 'Sum of month' => khởi chạy macros gatherData
*/

function gatherData() {
  var ss = SpreadsheetApp.getActive();
  var sumOfMonth = ss.getActiveSheet();

  var lastMer = sumOfMonth.getRange(2,26).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  var numOfMer = lastMer - 1;
  var merchants = sumOfMonth.getRange(2,26,numOfMer,1).getValues();
  
  for(var i = 0; i < numOfMer; i++){
    var lastRow = sumOfMonth.getLastRow();
    var start = lastRow + 1;
    var cell = sumOfMonth.getRange(start, 1);
    var merSheet = ss.getSheetByName(`Summary-${merchants[i]}`);
    var numRow = merSheet.getLastRow();
    merSheet.getRange(`A2:Y${numRow}`).copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  };

  var range = sumOfMonth.getRange(2,1,numOfMer,25);
  range.deleteCells(SpreadsheetApp.Dimension.ROWS);
};

