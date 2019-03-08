//this script is attached to spreadsheet BatchList
//trigger is event-based
//function updateForm is called On change
//function onOpen is called On Open

//create Custom Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Batch List', 'toBatchList')
      .addItem('Upload', 'toUpload')
      .addSeparator()
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}


//generate pop-up that will redirect user to spreadsheet BatchList
function toBatchList() {
  var selection = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  var html = "<script>window.open('https://docs.google.com/spreadsheets/d/1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM/edit?usp=sharing');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Batch List');
}

//generate pop-up that will redirect user to Google form Upload
function toUpload() {
  var selection = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  var html = "<script>window.open('https://goo.gl/forms/413VfKm8v6TR8zLz1');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Upload Metadata');
}

//update google form every time Batch List is edited
function updateForm(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("1o67QUW-I3WNYtHUBUNvTRk1T1gMW5OtXg9LNfVPqcIw");
   
  var namesList = form.getItemById("1059669708").asListItem();

// identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.getActive();
  var names = ss.getSheetByName("Batch List");

  // grab the values in the first column of the sheet - use 2 to skip header row 
  var namesValues = names.getRange(2, 1, names.getMaxRows() - 1).getValues();

  var batchName = [];

  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)    
    if(namesValues[i][0] != "")
      batchName[i] = namesValues[i][0];
  
  //add new batch option
  batchName.push("New Batch");

  // populate the drop-down with the array data
  namesList.setChoiceValues(batchName);
  
}
