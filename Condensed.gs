/**
 * @OnlyCurrentDoc
 */

function onOpen(e) {
  
  Logger.log(e.authMode);
  
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Submit Forms')
      .setWidth(300);
  var ui = SpreadsheetApp.getUi(); // Or DocumentApp or SlidesApp or FormApp.
  ui.showSidebar(html);
  ui.createMenu('Custom Menu')
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

function markSubmitted(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var range = sheet.getActiveRange();
  sheet.getRange(range.getRow(), 1, range.getNumRows(), 15).setBackground('#d9ead3');
  sheet.getRange(range.getRow(), 5, range.getNumRows(), 1).setNote('Submitted');
}

function markUnsubmitted(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var range = sheet.getActiveRange();
  sheet.getRange(range.getRow(), 1, range.getNumRows(), 15).setBackground('#ffcccb');
  sheet.getRange(range.getRow(), 5, range.getNumRows(), 1).setNote('Not yet submitted');
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
      sendEmail(temp,'NED NTP Receipt');
    }
    arr.push(temp);
  }
  
  arr = arr.reverse();
  Logger.log(arr);
  sheet.getRange(4, 1, arr.length, arr[0].length).setBackground('#d9ead3').setValues(arr);
  sheet.getRange(4, 5, arr.length, 1).setNote('Submitted');
}

function onNTPFormSubmit(e){
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
  var info = e.range.getValues()[0];
  info.shift(); //removes timestamp
  var ntp = info.shift(); //shifts out the Ntp number
  
  Logger.log(sheetName);
  Logger.log(ntp);
  Logger.log(info);
  
  updateReceiptByNTP(ntp,getUpdateCol(sheetName),info);  
  sendEmail(getNTPInfo(ntp),sheetName);
}

function getUpdateCol(sheetName){
  switch(sheetName){
    case 'NTP Assignment':
      return 16;
      break;
    case 'GIS Response':
      return 17;
      break;
    case 'Engineering Response':
      return 21;
      break;
    case 'GIS GDB Response':
      return 23;
      break;
    case 'VZ/LCS':
      return 25;
      break;
    default:
      return 1;
      break;
  }
}

function updateReceiptByNTP(ntp, startCol, insert){
 
  var numCol = insert.length;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  //get range from start to end of insert
  var range = sheet.getRange(4, 1, (sheet.getLastRow()-3), (startCol+numCol-1));
  var values = range.getValues();
  var flag = false;
  
  //iterate through sheet 
  for(var i = values.length-1; i >= 0; i--){
    
    //match NTPs
    if(values[i][4] == ntp){
      
      //inserting insert array
      for(var j = 0; j < numCol; j++){
        values[i][startCol-1+j] = insert[j];
      }
      
      //NTPs are grouped, once the last is entered, break the loop.
      flag = true;
    }else if(flag){
      break;
    }
  }
  
  range.setValues(values);
}

function sendEmail(ntpInfo, sheetName){
  
  var ntp = ntpInfo[4];
  var eI = emailInfo(ntp, sheetName);
  var body = getInfoHeader(ntpInfo);

  var tmpl = HtmlService.createTemplateFromFile(eI[2]);
  tmpl.ntp = ntp;
  body += tmpl.evaluate().getContent();
  
  MailApp.sendEmail({
    to: eI[0],
    subject: eI[1],
    htmlBody: body,
  });
}

function emailInfo(ntp, sheetName){
  
  switch(sheetName){
      
    case "NED NTP Receipt":
      //email going to Manager to assign an NTP to a tech
      return ['joseph.dyer@engineeringassociates.com',"To Assign: "+ntp,'ntp Assignment Template.html'];
      break;
      
    case 'NTP Assignment':
      //email going to tech to notify assignment
      return ['joseph.dyer@engineeringassociates.com',"You've been assigned "+ntp,'ntp assigned.html'];
      break;
      
    case 'GIS Response':
      //email going to engineering to notify completed GIS assignment
      return ['joseph.dyer@engineeringassociates.com',"GIS has finished NTP "+ntp, 'engineering form.html'];
      break;
      
    case 'Engineering Response':
      //email going to tech for GDB
      //find to whom NTP was assigned
      var AssignForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NTP Assignment');
      var AssignValues = AssignForm.getSheetValues(1, 1, AssignForm.getLastRow(), 3);
      var assignedTech;
      for(var i = AssignValues.length-1; i >=0; i--){
        if(AssignValues[i][1] == ntp){
          assignedTech = AssignValues[i][2].split(' ')[2];
          break;
        }
      }
      Logger.log(assignedTech);
      return [assignedTech, "Engineering has finished with "+ntp, 'GIS GDB.html'];
      break;
      
    case 'GIS GDB Response':
      //email to administration to notify GDB upload by GIS team
      return ['joseph.dyer@engineeringassociates.com',"GIS GDB has been submitted for "+ntp, 'VZ-LCS.html'];
      break;
      
    default:
      return;
      break;
  }
}

function getNTPInfo(ntp){
  
  var NTPForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var NTPValues = NTPForm.getSheetValues(1, 1, NTPForm.getLastRow(), 15);
  
  for(var i = NTPValues.length-1; i >= 0; i--){
    if(NTPValues[i][4] == ntp){
      return NTPValues[i];
      break;
    }
  } 
}

function getInfoHeader(NTPinfo){  
  return "<ul><li><b>Recept Date:</b> "+NTPinfo[1]+"</li><li><b>Market:</b> "+NTPinfo[2]+"</li><li><b>Shareable Link:</b> "+NTPinfo[3]+"</li><li><b>NTP Number:</b> "+NTPinfo[4]+"</li><li><b>VZW POR Number:</b> "+NTPinfo[5]+"</li><li><b>Include Customer Drop:</b> "+NTPinfo[6]+"</li><li><b>Wireless Location:</b> "+NTPinfo[7]+"</li><li><b>A LOC Access Point Description:</b> "+NTPinfo[8]+"</li><li><b>A LOC Access Point Lat:</b> "+NTPinfo[9]+"</li><li><b>A LOC Access Point Long:</b> "+NTPinfo[10]+"</li><li><b>Additional Fiber Length for Splice Point:</b> "+NTPinfo[11]+"</li><li><b>Z LOC Lat:</b> "+NTPinfo[12]+"</li><li><b>Z LOC Long:</b> "+NTPinfo[13]+"</li><li><b>NTP Work Description:</b> "+NTPinfo[14]+"</li></ul><br>";
}

function test(){
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  sheet.getRange(4, 1, sheet.getLastRow(), 15).setNote(null);
}