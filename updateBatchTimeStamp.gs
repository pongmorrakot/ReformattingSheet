//Trigger is time-based;

var batchListId = 'id of spreadsheets BatchList here';
var batchList;

//update value in Last Edit column in Batch List
function updateBatchTimeStamp() {
  var folder = DriveApp.getFoldersByName('Batches').next();
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var sheetId = file.getId();
    //var sheet = SpreadsheetApp.openById(sheetId);
    var lastEdit = file.getLastUpdated();
    var target = searchBatchList(sheetId);
    if(target != -1){
      var cell = batchList.getRange(target, 4);
      cell.setValue(lastEdit);
    }
  }
}

//search the batch to get the cell to update lastEdit to
//part of the code is adopted from: https://ctrlq.org/code/20001-find-rows-in-spreadsheets

function searchBatchList(id){
  batchList = SpreadsheetApp.openById(batchListId).getSheets()[0];
  var column = batchList.getRange('E:E');
  var values = column.getValues(); 
  var row = 0;
  while ( values[row] != id && values[row] != "") {
    row++;
  }
  if (values[row] == id) 
    return row+1;
  else 
    return -1;
}