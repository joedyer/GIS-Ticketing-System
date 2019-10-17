//function onNTPFormSubmit(e){
//  var sheet = e.source.getActiveSheet();
//  if(sheet.getName() == "NTP Form Responses"){
//    
//    //send out notification email
//    var arr = e.range.getValues();
//    sendAssignerEmail(arr);
//    
//    //set overview sheet
//    var NtpSheet = e.source.getSheetByName("NED NTP Receipt");
//    var por = arr[0][5];
//    arr[0][5] = 1;
//    
//    //add entries for as many POR as NTP has
//    for(var i = 2; i <= por; i++){
//      arr.push(arr[0].slice());
//      arr[i-1][5] = i;
//    }
//    
//    //set range
//    var range = NtpSheet.getRange(NtpSheet.getLastRow()+1, 1, arr.length, 15);
//    range.setValues(arr);
//    
//  } else if (sheet.getName() == "NTP Assignment") {
//    
//    var info = e.range.getValues();
//    //send out notification email to person assigned
//    sendAssignedEmails(info);
//    
//    //add name to assigned column
//    updateReceiptByNTP(info[0][2], 16, [info[0][1]]);
//    
//    
//  } else if (sheet.getName() == 'GIS Response') {
//    
//    var info = e.range.getValues()[0];
//    info.shift(); //removes timestamp
//    var ntp = info.shift(); //shifts out the Ntp number
//    
//    updateReceiptByNTP(ntp,17,info);
//    
//    sendEngineeringEmail(ntp);
// 
//  } else if (sheet.getName() == 'Engineering Response'){
//  
//    var info = e.range.getValues()[0];
//    info.shift(); //removes timestamp
//    var ntp = info.shift(); //shifts out the Ntp number
//    
//    updateReceiptByNTP(ntp,21,info);
//    
//    sendGISGDBEmail(ntp);
//  
//  } else if (sheet.getName() == 'GIS GDB Response'){
//    
//    var info = e.range.getValues()[0];
//    info.shift(); //removes timestamp
//    var ntp = info.shift(); //shifts out the Ntp number
//    
//    updateReceiptByNTP(ntp,23,info);
//  
//    sendAdminVZEmail(ntp);
//  }
//}
//
//function updateReceiptByNTP(ntp, startCol, insert){
// 
//  var numCol = insert.length;
//  var sheet = SpreadsheetApp.getActive().getSheetByName('NED NTP Receipt');
//  //get range from start to end of insert
//  var range = sheet.getRange(4, 1, (sheet.getLastRow()-3), (startCol+numCol-1));
//  var values = range.getValues();
//  var flag = false;
//  
//  //iterate through sheet 
//  for(var i = values.length-1; i >= 0; i--){
//    
//    //match NTPs
//    if(values[i][4] == ntp){
//      
//      //inserting insert array
//      for(var j = 0; j < numCol; j++){
//        values[i][startCol-1+j] = insert[j];
//      }
//      
//      //NTPs are grouped, once the last is entered, break the loop.
//      flag = true;
//    }else if(flag){
//      break;
//    }
//  }
//  
//  range.setValues(values);
//}
//
//
//function sendAssignerEmail(info){
//  
//  var ntp = info[0][4];
//  var subject = "To Assign: "+ntp+" in "+info[0][2];
//  
//  var body = getInfoHeader(info[0]);
//  var tmpl = HtmlService.createTemplateFromFile('ntp Assignment Template.html');
//  tmpl.ntp = ntp;
//  body += tmpl.evaluate().getContent();
//
//  MailApp.sendEmail({
//    to:"joseph.dyer@engineeringassociates.com",
//    subject:subject,
//    htmlBody: body,
//  });
//}
//
//function sendAssignedEmails(info){
//
//  var NTPinfo = getNTPInfo(info[0][2]);
//  
//  var subject = "You've been assigned "+NTP+" in "+NTPinfo[2];
//  var body = getInfoHeader(NTPinfo);
//  var tmpl = HtmlService.createTemplateFromFile('ntp assigned.html');
//  tmpl.ntp = NTP;
//  tmpl.market = NTPinfo[2];
//  body += tmpl.evaluate().getContent();
//
//  MailApp.sendEmail({
//    to: info[0][1].split(" ")[2],
//    subject:subject,
//    htmlBody: body,
//  });
//}
//
//function sendEngineeringEmail(ntp){
//  
//  var subject = "GIS has finished NTP "+ntp;
//  var ntpInfo = getNTPInfo(ntp);
//  var body = getInfoHeader(ntpInfo);
//  
//  var tmpl = HtmlService.createTemplateFromFile('engineering form.html');
//  tmpl.ntp = ntp;
//  body += tmpl.evaluate().getContent();
//
//  MailApp.sendEmail({
//    to: "joseph.dyer@engineeringassociates.com",
//    subject:subject,
//    htmlBody: body,
//  });
//}
//
//function sendGISGDBEmail(ntp){
// 
//  var subject = "Engineering has finished with "+ntp;
//  var ntpInfo = getNTPInfo(ntp);
//  var body = getInfoHeader(ntpInfo);
//  
//  var tmpl = HtmlService.createTemplateFromFile('GIS GDB.html');
//  tmpl.ntp = ntp;
//  body += tmpl.evaluate().getContent();
//  
//  //find to whom NTP was assigned
//  var AssignForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NTP Assignment');
//  var AssignValues = AssignForm.getSheetValues(1, 1, AssignForm.getLastRow(), 3);
//  var AssignedTech;
//  
//  for(var i = AssignValues.length-1; i >=0; i--){
//    if(AssignValues[i][2] == ntp){
//      AssignedTech = AssignValues[i][2].split(' ')[2];
//      break;
//    }
//  }
//  
//  MailApp.sendEmail({
//    to: AssignedTech,
//    subject:subject,
//    htmlBody: body,
//  });
//  
//}
//
//function sendAdminVZEmail(ntp){
//  
//  var subject = "GIS GDB has been submitted for "+ntp;
//  var ntpInfo = getNTPInfo(ntp);
//  var body = getInfoHeader(ntpInfo);
//  
//  var tmpl = HtmlService.createTemplateFromFile('engineering form.html');
//  tmpl.ntp = ntp;
//  body += tmpl.evaluate().getContent();
//
//  MailApp.sendEmail({
//    to: "joseph.dyer@engineeringassociates.com",
//    subject:subject,
//    htmlBody: body,
//  });
//  
//}
//
//function getNTPInfo(ntp){
//  
//  var NTPForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NTP Form Responses');
//  var NTPValues = NTPForm.getSheetValues(1, 1, NTPForm.getLastRow(), NTPForm.getLastColumn());
//  
//  for(var i = NTPValues.length-1; i >= 0; i--){
//    if(NTPValues[i][4] == ntp){
//      return NTPValues[i];
//      break;
//    }
//  } 
//  
//}
//
//function getInfoHeader(NTPinfo){
//  
//  return "<ul><li><b>Recept Date:</b> "+NTPinfo[1]+"</li><li><b>Market:</b> "+NTPinfo[2]+"</li><li><b>Shareable Link:</b> "+NTPinfo[3]+"</li><li><b>NTP Number:</b> "+NTPinfo[4]+"</li><li><b>VZW POR Number:</b> "+NTPinfo[5]+"</li><li><b>Include Customer Drop:</b> "+NTPinfo[6]+"</li><li><b>Wireless Location:</b> "+NTPinfo[7]+"</li><li><b>A LOC Access Point Description:</b> "+NTPinfo[8]+"</li><li><b>A LOC Access Point Lat:</b> "+NTPinfo[9]+"</li><li><b>A LOC Access Point Long:</b> "+NTPinfo[10]+"</li><li><b>Additional Fiber Length for Splice Point:</b> "+NTPinfo[11]+"</li><li><b>Z LOC Lat:</b> "+NTPinfo[12]+"</li><li><b>Z LOC Long:</b> "+NTPinfo[13]+"</li><li><b>NTP Work Description:</b> "+NTPinfo[14]+"</li></ul><br>";
//  
//}