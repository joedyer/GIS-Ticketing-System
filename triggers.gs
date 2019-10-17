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


function test(){
  
  var ntp = "dsafdsa";
  
  var tmpl = HtmlService.createTemplateFromFile('Copy of GIS Assignment.html');
  tmpl.ntp = ntp;
  var body = tmpl.evaluate().getContent();
  
  MailApp.sendEmail({
    to: 'joseph.dyer@engineeringassociates.com',
    subject: 'TESTING TEMPLATING',
    htmlBody: body,
  });
}
