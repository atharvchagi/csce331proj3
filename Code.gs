//https://developers.google.com/apps-script/guides/dialogs
//function onOpen() {
//  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
//    .createMenu("Custom Menu")
//    .addItem("Show alert", "showAlert")
//    .addToUi();
//}

function onOpen() {
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Stock Plotter')
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('My custom sidebar');
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


function generateChart(stockTicker, dateRange1, dateRange2, resolution, statistics, visualization) {
    //showAlert();
    //google.script.run.processStockData(stock, dateRange);
    //console.log("Server-side function called!");

    //debugging purposes
    Logger.log("stockTicker: " + stockTicker);
    Logger.log("dateRange1: " + dateRange1);
    Logger.log("dateRange2: " + dateRange2);
    Logger.log("resolution: " + resolution);
    Logger.log("statistics: " + statistics);
    Logger.log("visualization: " + visualization);

    //make the sheet and return the sheet so that it can be opened
    var ss = createStockSpreadsheet(stockTicker); //the spreadsheet that may have multiple pages
    var sheet = ss.getSheets()[0]; //the spreadsheet itself

    //do some last second formatting
    chartTitle = stockTicker;
    stockTicker = '"' + stockTicker + '"';
    attribute = '"close"';//FIXME, add attributes to GUI later
    dates1 = dateRange1.split("/");//starting date tokenized
    dates2 = dateRange2.split("/");//ending date tokenized

    //make google sheets use google finance formula to get stock data
    var cell = sheet.getRange('A1')
    cell.setValue("=GOOGLEFINANCE(" + stockTicker + "," + attribute + "," + "DATE(" + dates1[2] + "," + dates1[0] + "," + dates1[1] + ")" + "," + "DATE(" + dates2[2] + "," + dates2[0] + "," + dates2[1]+ "))");



    //create chart from this data
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange('A:B'))
      .setPosition(5, 5, 0, 0)
      .setOption('width', 480)
      .setOption('height', 300)
      .setOption('title', chartTitle)
      .setOption('vAxis.title', "Share Price")
      .setOption('hAxis.title', "Date")
      .setOption('hAxis', {slantedText: true, slantedTextAngle: 45})
      .build();

    
    
    sheet.insertChart(chart);
    

    //insert chart into google document
    //https://www.youtube.com/watch?v=ykFl0SE0rJE
    var doc = DocumentApp.getActiveDocument();
    SpreadsheetApp.flush();
    var chart_image = chart.getAs('image/png');
    var doc_body = doc.getBody();
    //doc.insertChart(chart);
    var inserted_image = doc_body.insertImage(1,chart_image);

    spreadsheetUrl = ss.getUrl();
    doc_body.appendParagraph('Here is the link to the Google Spreadsheet: ' + spreadsheetUrl);


    inserted_image.setLinkUrl(spreadsheetUrl);
    
}

function createStockSpreadsheet(stockTicker) {


  // Get the current date and time
  var currentDate = new Date();
  
  // Format the date and time to string (e.g., "yyyy-MM-dd HH:mm")
  var formattedDateTime = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  Logger.log(formattedDateTime);

  var sheetname = "StockSheet_" + stockTicker + "_" + formattedDateTime;


  var ss = SpreadsheetApp.create(sheetname);
  //debugging purposes
  Logger.log("created spreadsheet: " + sheetname);

  return ss;//return sheet so program can use

  
}




function showAlert() {
  var ui = DocumentApp.getUi(); // Same variations.

  var result = ui.alert(
    "Please confirm",
    "Are you sure you want to continue?",
    ui.ButtonSet.YES_NO,
  );

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert("Confirmation received.");
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Permission denied.");
  }
}