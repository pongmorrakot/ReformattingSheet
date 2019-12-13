var folderId = '1VQu1dlKgUx9c6iCUo-BqzaNQaY9oru12';
var archiveId = '1a8auoTmy0YoOHIvY7NyUf2rAore670_z';
var errorLogId = '1_CQ1nsOkKQXLbAR1rfeHgq4Js7gOtaaD3Zb6krVHqh4';
var batchListId = '1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM';
var archiveListName = 'Archived';

//when user delete an entry from the BatchList, a new entry will be added to the Archived page and the corresponding file moved to the Archived folder
//source: https://developers.google.com/drive/api/v3/folder#moving_files_between_folders

function archive() {
  var files = DriveApp.getFolderById(folderId).getFiles();
  var sheet = SpreadsheetApp.openById(batchListId).getSheetByName("Batch List")
  var data = sheet.getRange(2, 4, sheet.getLastRow()).getValues();//need to get how many row the batchList has
  data = flatten(data);
  //loop through all elements in the folder
  while(files.hasNext()){
    var file = files.next();
    var id = file.getId();
    var filename = file.getName();
    if(data.indexOf(id) == -1){
      var toMove = DriveApp.getFileById(id);
      var collectionId = SpreadsheetApp.openById(id).getSheetByName("Issue List").getRange(1, 2).getValue();
      addToArchiveList(toMove.getUrl(), filename, collectionId, id);
      DriveApp.getFolderById(archiveId).addFile(toMove);
      DriveApp.getFolderById(folderId).removeFile(toMove);      
      addLog(filename + ": archived");
    }
  }
}

function addToArchiveList(url, batchName, CollectionId, batchId){
  var archiveList = SpreadsheetApp.openById(batchListId).getSheetByName(archiveListName);
  archiveList.insertRowAfter(1);
  archiveList.getRange(2, 1, 1, 4).setValues([['=HYPERLINK(\"' + url + '\",\"' + batchName + '\")',
                                               CollectionId,
                                               new Date(),
                                               batchId]])
}

function flatten(array){
  return [].concat.apply([], array);;
}

function addLog(msg){
  var log = SpreadsheetApp.openById(errorLogId).getActiveSheet();
  log.insertRowAfter(1);
  log.getRange(2, 1, 1, 2).setValues([[new Date(),msg]])
}
