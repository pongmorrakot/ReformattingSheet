//this script is attached to UploadInfo
//the trigger is event-based
//this script is supposed to run every time there are changes in UploadInfo(user fill the Form Upload)
//more question can be asked in the Form to locate the information need from the spreadsheet uploaded
//new sheet is added to an existing batch if user chose to Add to an existing Batch
//new Batch spreadsheet is created if user chose to Add to New Batch
//an entries is added to BatchList containing BatchName, Collection Info, Date created, Date Last edited, and Batch spreadsheet's object ID
//this script is designed to generate at most one batch each time user submit a form
//this script could also be implemented to delete uploaded spreadsheet after data is extracted

//Drive API need to be enabled

//not yet debugged
//TODO: 

var batchListId = '1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM';
var batchList;
var collectionId;
var issueId;
var upload;//uploaded spreadshet file
var batchName;//contains the name of the Batch to add to; it can either be an existing batch or a new name
var batchId;

function onEdit(e ){
  //read info from this sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var i=0;
  while(data[i][0] != ""){
    var j=1;
    while(data[j][6] == "yes")
      j++;
    collectionId = data[j][1];
    issueId = data[j][2];
    upload = getPage(data[j][3]);
    var temp = data[j][4];
    if(temp == "New Batch")
      batchName = data[j][5];
    else
      batchName = temp;
    
    //decide to generate New Batch or add to an existing Batch
    if(batchName == "New Batch"){
      //generate New Batch
    }else{
      //add to an existing Batch
    }
    //mark Added column to yes
    i++;
  }
  
}

function createNewBatch(){
  //TODO: create new spreadsheet with the batchName as the name
  //first sheet list all of the issues in the batch with checklists
  //each issue is added as a sheet in the spreadsheet
  //Example: https://docs.google.com/spreadsheets/d/15TktmxhiY_0hxekAb75jrWeT36PtLUjbZ6d8UmRykws/edit?usp=sharing
}

function addToBatch(){
  //TODO: add new sheet to the spreadsheet with name matches with batchName
  //add a new row to IssueList sheet in the spreadsheet
}

//takes in url to excel file uploaded and return google sheet generated from the input
//generated google sheet is stored in temp folder
function getPage(url){
    var id = url.substring(33);
    var excelFile = DriveApp.getFileById(id);
    var blob = excelFile.getBlob();
    var folder = DriveApp.getFoldersByName("temp").next();
    var folderId = folder.getId();
    var resource = {
      title: excelFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: folderId}],
    };
    
    Drive.Files.insert(resource, blob);
    var files = folder.getFilesByName(excelFile.getName())
    var file = files.next()
    var ssId = file.getId();
    return SpreadsheetApp.openById(ssId);
}