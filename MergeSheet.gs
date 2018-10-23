/*Merge sheet tool*/
/*Any questions? -> ychen1987611@gmail.com */

function mergeSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var destiny = spreadsheet.insertSheet("all");
  var basesheet = spreadsheet.getSheetByName("common");
  var columnnum = basesheet.getMaxColumns();
  
  var allsheets = spreadsheet.getSheets();
  var beginrow = 1;
  for(var y in allsheets) 
  {
    if(allsheets[y].getMaxColumns() == columnnum)
    {
      var subsheetrow = allsheets[y].getMaxRows();
      
      if(allsheets[y].getName() != "common")
      {
        var data = allsheets[y].getRange(2, 1, subsheetrow - 1, columnnum);
        data.copyValuesToRange(destiny, 1, columnnum, beginrow, beginrow + subsheetrow - 2);
        beginrow += subsheetrow - 2;
      }
      else
      {
        var data = allsheets[y].getRange(1, 1, subsheetrow, columnnum);
        data.copyValuesToRange(destiny, 1, columnnum, beginrow, beginrow + subsheetrow - 1);
        beginrow += subsheetrow-1;
      }
    }
  }    
}

function CurrentSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function SheetNames() { // Usage as custom function: =SheetNames( GoogleClock() )
try {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  var out = new Array( sheets.length+1 ) ;
  //out[0] = [ "Name" , "gid" ];
  for (var i = 2 ; i < sheets.length+1 ; i++ ) out[i] = [sheets[i-1].getName()];
  return out
}
catch( err ) {
  return "#ERROR!" 
}
}

function addColumnOnAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var colIdxStr = Browser.inputBox("Insert column right of column no");
  if (colIdxStr == 'cancel') {
    return;
  }
  var colIdx = parseInt(colIdxStr);
  if (colIdx == NaN) {
    Browser.msgBox("You must enter a number or 'cancel'");
    return;
  }
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    sheet.insertColumnAfter(colIdx);
  }
}

function removeColumnOnAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var colIdxStr = Browser.inputBox("Remove column no");
  if (colIdxStr == 'cancel') {
    return;
  }
  var colIdx = parseInt(colIdxStr);
  if (colIdx == NaN) {
    Browser.msgBox("You must enter a number or 'cancel'");
    return;
  }
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    sheet.deleteColumn(colIdx);
  }
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {name: "Add column on all sheets", functionName: "addColumnOnAllSheets"}, 
    {name: 'Remove column on all sheets', functionName: 'removeColumnOnAllSheets'},
    {name: 'MergeSheets', functionName: 'mergeSheets'}
  ];
  sheet.addMenu("Columns", entries);
}
