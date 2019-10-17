function onNTPFormSubmit(e){
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
  var info = e.range.getValues()[0];
  info.shift(); //removes timestamp
  var ntp = info.shift(); //shifts out the Ntp number
  
  updateReceiptByNTP(ntp,getUpdateCol(sheetName),info);  
  sendEmail(ntp,sheetName, info[0]);
}

function getUpdateCol(sheetName){
  switch(sheetName){
    case 'GIS Assignment':
      return 16;
      break;
    case 'GIS Response':
      return 17;
      break;
    case 'Engineering Assignment':
      return 21;
      break;
    case 'Engineering Response':
      return 22;
      break;
    case 'GIS GDB Assignment':
      return 24;
      break;
    case 'GIS GDB Response':
      return 25;
      break;
    case 'VZ/LCS':
      return 27;
      break;
    default:
      return 1;
      break;
  }
}

function updateReceiptByNTP(ntp, startCol, insert){
 
  var numCol = insert.length;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var range = sheet.getRange(4, 1, (sheet.getLastRow()-3), (startCol+numCol-1));  //get range from start to end of insert
  var values = range.getValues();
  var flag = false;
  
  for(var i = values.length-1; i >= 0; i--){
    if(values[i][4] == ntp){ //match NTPs
      for(var j = 0; j < numCol; j++){
        values[i][startCol-1+j] = insert[j]; //inserting insert array
      }
      flag = true; //NTPs are grouped, once the last is entered, break the loop.
    }else if(flag){
      break;
    }
  }  
  range.setValues(values);
}

function sendEmail(ntp, sheetName, possibleAssign){
  
  var eI = emailInfo(ntp, sheetName, possibleAssign);
  var body = getInfoHeader(ntp);
  
  if (eI != null){
    var tmpl = HtmlService.createTemplateFromFile(eI[2]);
    tmpl.ntp = ntp;
    body += tmpl.evaluate().getContent();
    
    MailApp.sendEmail({
      to: eI[0],
      subject: eI[1],
      htmlBody: body,
    });
  }
}

function emailInfo(ntp, sheetName, assigned){
  
  switch(sheetName){
      
    case "NED NTP Receipt":
      //email going to Manager to assign an NTP to a tech
      return [SpreadsheetApp.getActive().getSheetByName('Teams').getRange(2, 4).getValue(),"To Assign: "+ntp,'GIS Assignment.html'];
      break;
      
    case 'GIS Assignment':
      //email going to tech to notify assignment
      return [assigned.split(' ')[2],"You've been assigned "+ntp,'GIS Assigned.html'];
      break;
      
    case 'GIS Response':
      //email going to engineering to notify completed GIS assignment
      return [SpreadsheetApp.getActive().getSheetByName('Teams').getRange(2, 8).getValue(),"GIS has finished NTP "+ntp, 'Engineering Assignment.html'];
      break;
      
    case 'Engineering Assignment':
      //email from engineering assigner to engineer assigned
      return [assigned.split(' ')[2],"You've been assigned "+ntp, 'Engineering Assigned.html'];
      break;
      
    case 'Engineering Response':
      //email going to GIS Assigner for GDB
      return [SpreadsheetApp.getActive().getSheetByName('Teams').getRange(2, 4).getValue(), "Engineering has finished with "+ntp, 'GIS GDB Assignment.html'];
      break;
      
    case 'GIS GDB Assignment':
      return [assigned.split(' ')[2],"You've been assigned "+ntp, 'GIS GDB Assigned.html'];
      break;
      
    case 'GIS GDB Response':
      //email to administration to notify GDB upload by GIS team
      return [SpreadsheetApp.getActive().getSheetByName('Teams').getRange(2, 10).getValue(),"GIS GDB has been submitted for "+ntp, 'VZ-LCS.html'];
      break;
      
    default:
      return null;
      break;
  }
}

function getInfoHeader(ntp){
  
  var NTPForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NED NTP Receipt');
  var NTPValues = NTPForm.getSheetValues(1, 1, NTPForm.getLastRow(), 15);
  var NTPinfo;
  
  for(var i = NTPValues.length-1; i >= 0; i--){
    if(NTPValues[i][4] == ntp){
      NTPinfo = NTPValues[i];
      break;
    }
  }
  return "<ul><li><b>Recept Date:</b> "+NTPinfo[1]+"</li><li><b>Market:</b> "+NTPinfo[2]+"</li><li><b>Shareable Link:</b> "+NTPinfo[3]+"</li><li><b>NTP Number:</b> "+NTPinfo[4]+"</li><li><b>VZW POR Number:</b> "+NTPinfo[5]+"</li><li><b>Include Customer Drop:</b> "+NTPinfo[6]+"</li><li><b>Wireless Location:</b> "+NTPinfo[7]+"</li><li><b>A LOC Access Point Description:</b> "+NTPinfo[8]+"</li><li><b>A LOC Access Point Lat:</b> "+NTPinfo[9]+"</li><li><b>A LOC Access Point Long:</b> "+NTPinfo[10]+"</li><li><b>Additional Fiber Length for Splice Point:</b> "+NTPinfo[11]+"</li><li><b>Z LOC Lat:</b> "+NTPinfo[12]+"</li><li><b>Z LOC Long:</b> "+NTPinfo[13]+"</li><li><b>NTP Work Description:</b> "+NTPinfo[14]+"</li></ul><br>";
}