var batchListId = '1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM';
var archivedFolderId = '1a8auoTmy0YoOHIvY7NyUf2rAore670_z';
var archiveListName = 'Archived';
  

function deleteBatch(){
//  delete all batches that is not on the batch list
//  get all files in the batch folder
//  get id for each file
//  compare it with id in the batch list 
//  delete those that is not in the batchlist

  var files = DriveApp.getFolderById(archivedFolderId).getFiles();
  var sheet = SpreadsheetApp.openById(batchListId).getSheetByName(archiveListName)
  var data = sheet.getRange(2, 4, sheet.getLastRow()).getValues();//need to get how many row the batchList has
  data = flatten(data);
  //loop through all elements in the folder
  while(files.hasNext()){
    var file = files.next();
    var id = file.getId();
    var filename = file.getName();
    if(data.indexOf(id) == -1){
      Drive.Files.remove(id);
    }
  }
}

function flatten(array){
  return [].concat.apply([], array);;
}
