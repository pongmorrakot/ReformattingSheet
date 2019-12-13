function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('UWDC')
      .addItem('Batch List', 'toBatchList')
      .addItem('Upload', 'toUpload')
      .addSeparator()
      .addItem('User Manual', 'toManual')
      .addItem('Log', 'toErrorLog')
      .addToUi();
}

function toBatchList() {
  var html = "<script>window.open('https://docs.google.com/spreadsheets/d/1ubmetfPvfJYJiug6XArwXgM_HILRlm-Ewfxe1tPsSnM/edit?usp=sharing');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Batch List');
}


function toUpload() {
  var html = "<script>window.open('https://docs.google.com/forms/d/e/1FAIpQLSfvo0g_u_E0wAHJ-4jodarWxvFV-17SpKUDN7eoaAaGZiVMWA/viewform?usp=pp_url&entry.583221579=New+Batch', '_blank');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Upload Metadata');
}

function toManual(){
  var html = "<script>window.open('https://docs.google.com/document/d/18rnI6rHGNh8R9d5KfBBCW6hqhmbzCc738X2AvazZbz4/edit?usp=sharing', '_blank');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'User Manual');
}

function toErrorLog(){
  var html = "<script>window.open('https://docs.google.com/spreadsheets/d/1_CQ1nsOkKQXLbAR1rfeHgq4Js7gOtaaD3Zb6krVHqh4/edit?usp=sharing', '_blank');google.script.host.close();</script>"
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Error Log');
}