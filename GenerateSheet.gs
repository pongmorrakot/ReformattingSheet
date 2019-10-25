//this script is attached to UploadInfo and is triggered everytime user upload using the form
//new sheet is added to an existing batch if user chose to Add to an existing Batch
//new Batch spreadsheet is created if user chose to Add to New Batch
//an entries is added to BatchList containing BatchName, Collection Info, Date created, Date Last edited, and Batch spreadsheet's object ID
//this script could also be implemented to delete uploaded spreadsheet after data is extracted

//Drive API need to be enabled

//useful document:
//Best practice: https://developers.google.com/apps-script/guides/support/best-practices

//Configuration

var debug = true;

var batchListId = '1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM';
var formId = "1o67QUW-I3WNYtHUBUNvTRk1T1gMW5OtXg9LNfVPqcIw";
var formDropDownId = "1059669708";
var folderId = '1VQu1dlKgUx9c6iCUo-BqzaNQaY9oru12';
var uploadInfoId = '1viCc9O4q7EpahumOsQvrN-n1nGlgZqhlBPfqUh0bdDg';
var errorLogId = '1_CQ1nsOkKQXLbAR1rfeHgq4Js7gOtaaD3Zb6krVHqh4';

var targetSheetName = "Page";
var issueListData = [["Issue","Visually Cohesive","Total Image Number","Number Sequence","Image Size","Photoshop Check","Thumbnail Check","Capture One","Bridge QC","Notes"]];
var issueData = [["File Name","Page No.","Page Notes","Scan","Crop","Notes"],
                 [    3,          4,         10,         -1,   -1,    -1   ]];

//Shouldn't need to change anything from here
//=================================================================================================================================================================================
//----------------------------------------------------------------------------------------Line of Abstraction--------------------------------------------------------------------------------
//=================================================================================================================================================================================


var collectionId;
var issueId;
var upload;//uploaded spreadsheet file
var batchName;//contains the name of the Batch to add to; it can either be an existing batch or a new name
var batchId;

var currentRow;

function onChange(e){
  //read info from this sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var batch;//contain the spreadsheet that new data is going to be added to
  var upload;
  var isNew = false;
  var url; // contains the url of newly construct batch spreadsheet; empty if not add to new batch
  
  var arrayLength = data.length;
  currentRow = 1;
  while(currentRow < arrayLength){
    if(data[currentRow][6] == ""){
      
      collectionId = data[currentRow][1];
      issues = data[currentRow][2].split("\n");
      issues = issues.filter(Boolean);
      upload = getPage(data[currentRow][3]);
      batchName = data[currentRow][4];
      batchId = batchName.split(' ')[batchName.split(' ').length - 1];
      if(batchName == "New Batch") isNew = true;
      
      for(var j=0; j < issues.length; j++){
        issueId = issues[j];
        if(isNew) url = initNewBatch(currentRow, data);// throw an error from this func
        
        batch = DriveApp.getFileById(batchId);
        
        batch = SpreadsheetApp.open(batch);
        
        var sheetId = addToBatch(upload,batch); //throw an error from this func
        addToList(sheetId,batch); //throw an error from this func
        
        if(isNew){
          addToBatchList(url);
          isNew = false;
        }
      }      
      
      //mark Added column to yes
      //indicating that import is done
      SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("yes");
      updateForm();
      SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),batchName + ' import successful']);
    }
    currentRow++;
  }
  
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
    var file = Drive.Files.insert(resource, blob);
    if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"getPage(): successful"]);
    return SpreadsheetApp.openById(file.id);
}

//create a new spreadsheet that would contain an Issue List
//and add the spreadsheet as an entry in the batch list
//possible error: cannot get batchList(i.e. got deleted, id doesn't match)
function initNewBatch(index, data){
  //create a sheet in a folder
  //https://stackoverflow.com/questions/19607559/how-to-create-a-spreadsheet-in-a-particular-folder-via-app-script
  batchName = data[index][5];
  var file = {
    title: batchName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  }
  file = Drive.Files.insert(file);
  batchId = file.id;
  var batch = SpreadsheetApp.openById(file.id);
  //add issue list to the new batch
  batch.getActiveSheet().getRange(1, 1, 1, issueListData[0].length).setValues(issueListData);
  batch.renameActiveSheet("Issue List");
  batch.getActiveSheet().autoResizeColumns(2, issueListData[0].length-1);
  batch.getActiveSheet().getRange(2, 2, 100, 6).setHorizontalAlignment("center");
  if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"initNewBatch(): successful"]);
  return url = batch.getUrl();
}

//select data from metadata spreadsheet and add it to the issue sheet
//return: id of the sheet if successful
//possible error: cannot find entries with the entered Issue_ID, upload, batch, issueId is not initialized
function addToBatch(upload, batch){
  var data = upload.getSheetByName(targetSheetName).getDataRange().getValues();
  var input = [];
  input[1] = issueData[0];
  var LocFound = false;
  var startLocator;
  var j = 1;
  var arrayLength = data.length;
  for(var i = 2; i < arrayLength; i++){
    if(data[i][1] == issueId){ 
      if(!LocFound){
        startLocator = i;
        LocFound = true;
      }
      //change this into a loop that iterate through issueData
      input[j] = [];
      for(var k = 0; k < issueData[1].length; k++){
        var num = this.issueData[1][k];
        if(num >= 0) input[j][k] = data[i][num];
        else input[j][k] = "";
      }
      j++;
    }
  }
  if(!LocFound){
    SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("skip");
    var msg = "addToBatch: no entry with given issueId: \"" + issueId + "\" is found"
    SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),msg]);
    throw msg;
  }
  var head = ["Page Location: " + data[startLocator][7]];
  for(var k = 1; k < issueData[1].length; k++) head.push("");
  input[0] = head;
  batch.insertSheet(issueId);
  var sheet = batch.getActiveSheet();
  sheet.getRange(1, 1, j, issueData[1].length).setValues(input);
  sheet.autoResizeColumns(2,issueData[1].length);
//  sheet.getRange(2, 2, j-1, issueData.length[1]-1).setHorizontalAlignment("center");
  if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"addToBatch(): successful"]);
  return sheet.getSheetId();
}

// add the new issue entry to the issue list
// possible error: cannot find page called issue List(i.e. Issue List got deleted), issueId is not initialized(shouldn't have reach this point)
function addToList(sheetId,batch){
  var sheet = batch.getSheetByName("Issue List");
  if(sheet == null){
    SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("skip");
    var msg = "addToList: Issue List not found";
    SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),msg]);
    throw msg;
  }
  var data = sheet.getDataRange().getValues();
  var done = false;
  var arrayLength = data.length;
  var i = 0
  while(i < arrayLength && !done){
    if(data[i][0] == ""){
      sheet.getRange(i+1, 1).setValue("=HYPERLINK(\"#gid=" + sheetId + "\",\"" + issueId + "\")");
      done = true;
    }
    i++
  }
  if(!done){
    sheet.getRange(i+1, 1).setValue("=HYPERLINK(\"#gid=" + sheetId + "\",\"" + issueId + "\")");
  }
  if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"addToList(): successful"]);
}

//add the batch sheet to the batch list
//TODO: add the file file id to the batchlist 
function addToBatchList(url){
  var batchList = SpreadsheetApp.openById(batchListId).getActiveSheet();
  if(batchList == null){
    SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("skip");
    var msg = " initNewBatch: Batch List not found"
    SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),msg]);
    throw msg;
  }
  var batchData = batchList.getDataRange().getValues();
  var done = false;
  var arrayLength = batchData.length;
  var i = 0
  var batchListEntry = [['=HYPERLINK(\"' + url + '\",\"' + batchName + '\")',
                         collectionId,
                         new Date(),
                         batchId]];
  while(i < arrayLength && !done){
    if(batchData[i][0] == ""){
      batchList.getRange(i+1, 1, 1, 4).setValues(batchListEntry);
      done = true;
    }
    i++
  }
  if(!done) batchList.getRange(arrayLength+1, 1, 1, 4).setValues(batchListEntry);
  if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"addToBatchList(): successful"]);
}



function updateForm() {
  // call your form and connect to the drop-down item
  var form = FormApp.openById(formId);
   
  var namesList = form.getItemById(formDropDownId).asListItem();

// identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.openById(batchListId);
  var names = ss.getSheetByName("Batch List");

  // grab the values in the first column of the sheet - use 2 to skip header row 
  var namesValues = names.getRange(2, 1, names.getMaxRows() - 1,4).getValues();
  

  var batchName = [];

  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++){
    if(namesValues[i][0] != ""){
      batchName[i] = namesValues[i][0] + " " + namesValues[i][3];
    }
}
  //add new batch option
  batchName.unshift("New Batch");

  // populate the drop-down with the array data
  namesList.setChoiceValues(batchName);
  if(debug) SpreadsheetApp.openById(errorLogId).getActiveSheet().appendRow([new Date(),"updateForm(): successful"]);
}
