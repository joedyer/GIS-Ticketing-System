function myFunction() {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('Charts');
  var range = sheet.getRange(1,1,9,2);
  range.setValues(getOpenAndClosedNTPs());
  
  var chart = sheet.newChart();
  
  chart = chart
          .asPieChart()
          .addRange(range)
          .setNumHeaders(1)
          .setPosition(1, 1, 0, 0)
          .setOption('animation.duration', 1000)
          .setColors(['#fff2cc','#d0e0e3','#76a5af','#ead1dc','#c27ba0','#c9daf8','#6d9eeb','#ffd966'])
          .set3D();
  
  sheet.insertChart(chart.build());
  
  var lastrow = sheet.getLastRow();
  
  var GISrange = sheet.getRange(22, 1, lastrow-22, 2);
  var ENGrange = sheet.getRange(22, 3, lastrow-22, 2);
  var GDBrange = sheet.getRange(22, 5, lastrow-22, 2);
  
  
  var GISchart = sheet.newChart();
  var Engchart = sheet.newChart();
  var GDBchart = sheet.newChart();
  
  
  GISchart = GISchart
          .asColumnChart()
          .addRange(GISrange)
          .setNumHeaders(1)
          .setPosition(1, 8, 0, 0)
          .setYAxisTitle('Assigned NTPs')
          .setTitle('GIS team assigned')
          .setOption('animation.duration', 1000);
  sheet.insertChart(GISchart.build());
  
    Engchart = Engchart
          .asColumnChart()
          .addRange(ENGrange)
          .setNumHeaders(1)
          .setPosition(21, 1, 0, 0)
          .setYAxisTitle('Assigned NTPs')
          .setTitle('Eng. team assigned')
          .setOption('animation.duration', 1000);
  sheet.insertChart(GISchart.build()); 
  
  GISchart = GISchart
          .asColumnChart()
          .addRange(GISrange)
          .setNumHeaders(1)
          .setPosition(21, 8, 0, 0)
          .setYAxisTitle('Assigned NTPs')
          .setTitle('GIS (GDB) team assigned')
          .setOption('animation.duration', 1000);
  sheet.insertChart(GISchart.build());
}

function clearCharts() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Charts');
  var charts = sheet.getCharts();
  
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
}

function updateCharts() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Charts');
  var range = sheet.getRange(1,1,9,2);
  range.setValues(getOpenAndClosedNTPs()); 
}

function getOpenAndClosedNTPs(){
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('NED NTP Receipt');
  var range = sheet.getRange(4, 1, sheet.getLastRow(), 27);
  var values  = range.getValues();
 
  var finalValues = [['Stages', 'Number of NTPs'],['Open NTPs'],['GIS Open'],['GIS Closed'],['Engineering Open'],['Engineering Closed'],['GIS GDB Open'],['GIS GDB Closed'],['Complete']];
  
  var colsToCheck =  [5, 16, 17, 21, 22, 24, 25, 27];
  
  for(var x = 0; x < colsToCheck.length; x++){
    var sum = 0;
    for(var i = 0; i < values.length; i++){
      if(values[i][colsToCheck[x]-1] != ''){
        sum++;
      }
    }
    finalValues[x+1].push(sum);
  }
  
  for(var i = 1; i < finalValues.length-1; i++){
    finalValues[i][1] = finalValues[i][1]-finalValues[i+1][1];
  }
    
 return finalValues;
}