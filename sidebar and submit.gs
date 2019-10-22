function addMenu() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .createMenu('Custom Menu')
        .addItem('Submit Forms', 'submitFromReceipt')
        .addItem('Open Sidebar', 'openSideBar')
        .addToUi();
}

function openSideBar(){
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Submit Forms')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function markSubmitted(color, note){
  if(color == null || note == null){
    color = '#ffcccb';
    note = 'Not yet submitted';
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var range = sheet.getActiveRange();
  var row = range.getRow();
  if (row > 3){
    sheet.getRange(row, 1, range.getNumRows(), 15).setBackground(color);
    sheet.getRange(row, 5, range.getNumRows(), 1).setNote(note);
  }
}

function submitFromReceipt(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var range = sheet.getRange(4, 1, sheet.getLastRow()-3, 15);
  var values = range.getValues();
  var bgs = range.getBackgrounds();
  var arr = [];
  var date = Utilities.formatDate(new Date(), 'EST', "MM/dd/yyy");
  
  for(var i = values.length-1; i >= 0; i--){
    var temp = values[i].slice();
    
    if(temp[5]==""){
      temp[5]='1';
    }
    
    if(bgs[i][0] == '#ffcccb'){
      temp[0] = date;
      for(temp[5]; temp[5] > 1; temp[5]--){
        arr.push(temp.slice());
      }
      sendEmail(temp[4],'NED NTP Receipt','');
    }
    arr.push(temp);
  }
  
  arr = arr.reverse();
  sheet.getRange(4, 1, arr.length, arr[0].length).setBackground('#d9ead3').setValues(arr);
  sheet.getRange(4, 5, arr.length, 1).setNote('Submitted');
}
