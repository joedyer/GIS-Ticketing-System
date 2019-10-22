function triggerSetup() {
  var ss = SpreadsheetApp.getActive();
  
  //delete all previously set up triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  
  ScriptApp.newTrigger("edit").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("open").forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger("onNTPFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
}

function edit(e){
  var sheet  = e.source.getActiveSheet();
  var name = sheet.getName();
  
  if(name == 'Teams'){
    refreshTeamOptions();
  }else if (name == 'NED NTP Receipt'){
    markSubmitted('#ffcccb', 'Not yet submitted');
  }
}

function open(){
  
  addMenu();
  openSideBar();
  
  //makesure hidden sheets are hidden
  var sheets = SpreadsheetApp.getActive().getSheets();
  
  for(var i = 0; i < sheets.length; i++){
    var sheetname = sheets[i].getName();
    if(sheetname != 'NED NTP Receipt' && sheetname != 'Charts' && sheetname != 'Data/Instructions' && sheetname != 'NED NTP' && sheetname != 'Teams'){
      sheets[i].hideSheet();
    }
  }
}