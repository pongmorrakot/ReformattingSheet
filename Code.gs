//this script is attached to UploadInfo and is triggered everytime user upload using the form
//new sheet is added to an existing batch if user chose to Add to an existing Batch
//new Batch spreadsheet is created if user chose to Add to New Batch
//an entries is added to BatchList containing BatchName, Collection Info, Date created, Date Last edited, and Batch spreadsheet's object ID
//this script could also be implemented to delete uploaded spreadsheet after data is extracted

//Drive API need to be enabled

//useful document:
//Best practice: https://developers.google.com/apps-script/guides/support/best-practices

//Configuration

//setting this to true can affect performance as it significantly increase number of write
var debug = false;

//id of a doc/sheet/form can be found in their url
//for example: https://docs.google.com/spreadsheets/d/1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM/edit
//                                                    ^..........................................^
//                               this document's id = 1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM
var batchListId = '1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM';
var formId = "1o67QUW-I3WNYtHUBUNvTRk1T1gMW5OtXg9LNfVPqcIw";
var formDropDownId = "1059669708";
var folderId = '1VQu1dlKgUx9c6iCUo-BqzaNQaY9oru12';
var uploadInfoId = '1viCc9O4q7EpahumOsQvrN-n1nGlgZqhlBPfqUh0bdDg';
var errorLogId = '1_CQ1nsOkKQXLbAR1rfeHgq4Js7gOtaaD3Zb6krVHqh4';

var targetSheetName = "Page"; //name of sheet in source excel spreadsheet to parse data from
var issueListData = [["Issue Id","Capture One","Total Image Number", "Total Entry Number", "Number Sequence","Image Size","Photoshop Check","Thumbnail Check","Visually Cohesive"],//column in issuelist that will be generated; entry will list under first column
                     [   0,           -1,              -1,                    1,                -1,              -1,              -1,              -1,               -1]];
//                       0 means the name of the issue added to this column   1 means that the total 
var issueData = [["File Name","Page No.","Page Notes","Scan","Crop"], //column in issue that will be generated
                 [    3,          4,         10,         -1,   -1]]; //indicate which column in source to parse from; -1 means leaving the column blank

// used in issueList to add conditional formatting rule 
// could be improved
function setRule(sheet, rowNum, col1, col2){
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberEqualTo(Number(sheet.getRange(rowNum, col1).getValue()))
  .setBackground("#9AE76D")
  .setRanges([sheet.getRange(rowNum, col2)])
  .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

//Shouldn't need to change anything from here
//=================================================================================================================================================================================
//----------------------------------------------------------------------------------------Line of Abstraction--------------------------------------------------------------------------------
//=================================================================================================================================================================================


var collectionId;
var issueId;
var upload;//uploaded spreadsheet file
var batchName;//contains the name of the Batch to add to; it can either be an existing batch or a new name
var batchId;

var currentRow;//current row in UploadInfo that is being processed

//write to log with timestamp
//msg: message to write
function addLog(msg){
  var log = SpreadsheetApp.openById(errorLogId).getActiveSheet();
  log.insertRowAfter(1);
  log.getRange(2, 1, 1, 2).setValues([[new Date(),msg]])
}

//main method of the script
//runs when user upload
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
      if(batchName == "New Batch"){
        isNew = true;
        batchName = data[currentRow][5];
      }
      addLog("Batch: " + batchName + ': import start');
      
      for(var j=0; j < issues.length; j++){
        issueId = issues[j];
        if(isNew) url = initNewBatch();// throw an error from this func
        
        batch = DriveApp.getFileById(batchId);
        
        batch = SpreadsheetApp.open(batch);
        
        var sheetId = addToBatch(upload,batch); //throw an error from this func
        if(isNew){
          addToBatchList(url);
          isNew = false;
        }
        addToList(sheetId,batch); //throw an error from this func
      }      
      
      //mark Added column to yes
      //indicating that import is done
      SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("yes");
      updateForm();
      addLog("Batch: " + batchName + ': import successful');
    }
    currentRow++;
  }
  
}

//takes in url to excel file uploaded and return google sheet generated from the input
//generated google sheet is stored in temp folder
//return: google sheet generated
//url: url to the excel file to turn into google sheet
function getPage(url){
  if(debug) addLog("getPage(): start");
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
    if(debug) addLog("getPage(): successful");
    return SpreadsheetApp.openById(file.id);
}

//create a new spreadsheet that would contain an Issue List
//return: url of spreadsheet created
function initNewBatch(){
  if(debug) addLog("initNewBatch(): start");
  //create a sheet in a folder
  //https://stackoverflow.com/questions/19607559/how-to-create-a-spreadsheet-in-a-particular-folder-via-app-script
  var file = {
    title: batchName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  }
  file = Drive.Files.insert(file);
  batchId = file.id;
  var batch = SpreadsheetApp.openById(file.id);
  var sheet = batch.getActiveSheet();
  //add issue list to the new batch
  sheet.getRange(1, 1, 1, 2).setValues([["Collection ID: ", collectionId]]);
  var input = [];
  input[0] = issueListData[0];
  sheet.getRange(2, 1, 1, issueListData[0].length).setValues(input);
  sheet.setName("Issue List");
  sheet.autoResizeColumns(2, issueListData[0].length-1);
  sheet.getRange(2, 2, 100, 6).setHorizontalAlignment("center");
  if(debug) addLog("initNewBatch(): successful");
  return url = batch.getUrl();
}

//select data from metadata spreadsheet and add it to the issue sheet
//return: id of the sheet if successful
//upload: Google sheet that contains uploaded data
//batch: spreadsheet to add issue to
function addToBatch(upload, batch){
  if(debug) addLog("addToBatch(): start");
  var data = upload.getSheetByName(targetSheetName).getDataRange().getValues();
  var input = [];
  input[1] = issueData[0];
  var LocFound = false;
  var startLocator;
  var j = 2;
  var arrayLength = data.length;
  for(var i = 1; i < arrayLength; i++){
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
    addLog(msg);
    throw msg;
  }
  var head = ["Page Location: " + data[startLocator][7]];
  for(var k = 1; k < issueData[1].length; k++) head.push("");
  input[0] = head;
  batch.insertSheet(issueId);
  var sheet = batch.getActiveSheet();
  sheet.getRange("A:B").setNumberFormat("@");
  sheet.getRange(1, 1, j, issueData[1].length).setValues(input);
  sheet.autoResizeColumns(2,issueData[1].length);
  sheet.setFrozenRows(2);
  if(debug) addLog("addToBatch(): successful");
  return sheet.getSheetId();
}

// add a new issue entry to the issue list
//sheetId: id of sheet(issue) to add to issuelist
//batch: spreadsheet that both issue list and issues are in 
function addToList(sheetId,batch){
  if(debug) addLog("addToList(): start");
  var sheet = batch.getSheetByName("Issue List");
  if(sheet == null){
    SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("skip");
    var msg = "addToList: Issue List not found";
    addLog(msg);
    throw msg;
  }
  
  var nameIndex;
  var numIndex;
  for(k = 0; k < issueListData[1].length; k++) if(issueListData[1][k] == 0) nameIndex = k;
  for(k = 0; k < issueListData[1].length; k++) if(issueListData[1][k] == 1) numIndex = k;
  var data = sheet.getDataRange().getValues();
  var done = false;
  var arrayLength = data.length;
  var i = 0
  while(i < arrayLength && !done){
    if(data[i][nameIndex] == ""){
      sheet.getRange(i+1, nameIndex+1).setValue("=HYPERLINK(\"#gid=" + sheetId + "\",\"" + issueId + "\")");
      sheet.getRange(i+1, numIndex+1).setValue(batch.getSheetByName(issueId).getDataRange().getValues().length - 2)
      setRule(sheet,i+1,numIndex+1,numIndex);
      done = true;
    }
    i++
  }
  if(!done){
    sheet.getRange(i+1, nameIndex+1).setValue("=HYPERLINK(\"#gid=" + sheetId + "\",\"" + issueId + "\")");
    sheet.getRange(i+1, numIndex+1).setValue(batch.getSheetByName(issueId).getDataRange().getValues().length - 2)
    setRule(sheet,i+1,numIndex+1,numIndex);
  }
  if(debug) addLog("addToList(): successful");
}

//add a batch to the batch list
//url: url of the batch to add to batch list
function addToBatchList(url){
  if(debug) addLog("addToBatchList(): start");
  var batchList = SpreadsheetApp.openById(batchListId).getActiveSheet();
  if(batchList == null){
    SpreadsheetApp.openById(uploadInfoId).getActiveSheet().getRange(currentRow+1,7).setValue("skip");
    var msg = " initNewBatch: Batch List not found"
    addLog(msg);
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
  batchList.insertRowAfter(1);
  batchList.getRange(2, 1, 1, 4).setValues(batchListEntry);
//  while(i < arrayLength && !done){
//    if(batchData[i][0] == ""){
//      batchList.getRange(i+1, 1, 1, 4).setValues(batchListEntry);
//      done = true;
//    }
//    i++
//  }
//  if(!done) batchList.getRange(arrayLength+1, 1, 1, 4).setValues(batchListEntry);
  if(debug) addLog("addToBatchList(): successful");
}

//update Google Form to contain every batch so user can add new issues to a batch
function updateForm() {
  if(debug) addLog("updateForm(): start");
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
  if(debug) addLog("updateForm(): successful");
}