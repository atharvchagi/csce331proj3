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
      .setTitle('Stock Plotter');
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


function generateChart(stockTicker, attribute, dateRange1, dateRange2, resolution, indicators, statistics, visualization) {
  //showAlert();
  //google.script.run.processStockData(stock, dateRange);
  //console.log("Server-side function called!");

  //debugging purposes
  //Logger.log("stockTicker: " + stockTicker);
  //Logger.log("dateRange1: " + dateRange1);
  //Logger.log("dateRange2: " + dateRange2);
  //Logger.log("resolution: " + resolution);
  //Logger.log("indicators: " + indicators);
  //Logger.log("visualization: " + visualization);

  var ui = DocumentApp.getUi();
  ui.alert("Creating New Visual Representation\n\n" + 
           "Stock Ticker: " + stockTicker + 
           "\nAttribute: " + attribute + 
           "\nStart Date: " + dateRange1 + 
           "\nEnd Date: " + dateRange2 + 
           "\nResolution: " + resolution + 
           "\nIndicators: " + indicators + 
           "\nStatistics: " + statistics + 
           "\nVisualization: " + visualization);

  //make the sheet and return the sheet so that it can be opened
  var ss = createStockSpreadsheet(stockTicker); //the spreadsheet that may have multiple pages
  var sheet = ss.getSheets()[0]; //the spreadsheet itself

  //do some last second formatting
  chartTitle = stockTicker;
  stockTicker = '"' + stockTicker + '"';
  resolution = '"' + resolution + '"';
  attribute_axis_name = attribute.charAt(0).toUpperCase() + attribute.slice(1).toLowerCase();
  attribute = '"' + attribute + '"';
  dates1 = dateRange1.split("/");//starting date tokenized
  dates2 = dateRange2.split("/");//ending date tokenized

  //make google sheets use google finance formula to get stock data
  var cell = sheet.getRange('A1')
  cell.setValue("=GOOGLEFINANCE(" + stockTicker + ","
  + attribute + "," 
  + "DATE(" + dates1[2] + "," + dates1[0] + "," + dates1[1] + ")," 
  + "DATE(" + dates2[2] + "," + dates2[0] + "," + dates2[1]+ "),"
  + resolution + ")");

  
  //create chart from this data
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('A:A'))
    .addRange(sheet.getRange('B:B'))
    .setPosition(5, 5, 0, 0)
    .setOption('width', 480)
    .setOption('height', 320)
    .setOption('title', chartTitle)
    .setOption('vAxis.title', attribute_axis_name)
    .setOption('hAxis.title', "Date")
    .setOption('hAxis', {slantedText: true, slantedTextAngle: 45})

  var series_labels = {
    0: { labelInLegend: 'Stock Price' }
  };
  var next_series_num = 1

  
  //add indicators/statistics before finalizing chart
  //20 day moving average
  if (indicators % 2 == 1) // using binary code to encode indicators
  {
    calcMovingAverage(sheet,3); //Calculate moving average in column C
    series_labels[next_series_num] = { labelInLegend: 'Moving Average' };
    next_series_num = next_series_num + 1;
    chart.addRange(sheet.getRange('C:C')); //moving average
  }
  //14 day RSI
  if ((indicators >>> 1) % 2 == 1)
  {
    calcRSI(sheet,4); //Calculate RSI, with sub calculations starting in column D, RSI will be in column J
    series_labels[next_series_num] = { labelInLegend: 'RSI' };
    next_series_num = next_series_num + 1;
    chart.addRange(sheet.getRange('J:J')); //RSI
  }
  //MACD (12,26) and Signal (9)
  if ((indicators >>> 2) % 2 == 1)
  {
    calcMACDandSignal(sheet,11); //Calculate MACD and signal. Subcalculations start in column K, MACD in Column M, signal in column N
    series_labels[next_series_num] = { labelInLegend: 'MACD' };
    next_series_num = next_series_num + 1;
    series_labels[next_series_num] = { labelInLegend: 'Signal' };
    next_series_num = next_series_num + 1;
    chart.addRange(sheet.getRange('M:M')); //MACD
    chart.addRange(sheet.getRange('N:N')); //Signal
  }


  chart.setOption('series', series_labels);

  //put chart into google sheets
  chart = chart.build();
  sheet.insertChart(chart);

  //insert chart into google document
  //https://www.youtube.com/watch?v=ykFl0SE0rJE
  var doc = DocumentApp.getActiveDocument();
  SpreadsheetApp.flush();
  if (visualization == "Chart")
  {
    var chart_image = chart.getAs('image/png');
    var doc_body = doc.getBody();
    //doc.insertChart(chart);
    var inserted_image = doc_body.insertImage(1,chart_image);

    spreadsheetUrl = ss.getUrl();
    //doc_body.appendParagraph('Here is the link to the Google Spreadsheet: ' + spreadsheetUrl);

    inserted_image.setLinkUrl(spreadsheetUrl); //put link to spreadsheet in image
      


    if ((indicators >>> 2) % 2 == 1) //also make separate plot for MACD
    {
      var macd_chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange('A:A'))
      .addRange(sheet.getRange('M:M')) //MACD
      .addRange(sheet.getRange('N:N')) //Signal
      .setPosition(5, 5, 0, 0)
      .setOption('width', 480)
      .setOption('height', 320)
      .setOption('title', chartTitle)
      .setOption('vAxis.title', "MACD and Signal")
      .setOption('hAxis.title', "Date")
      .setOption('hAxis', {slantedText: true, slantedTextAngle: 45})

      var macd_series_labels = {
        0: { labelInLegend: 'MACD' },
        1: { labelInLegend: 'Signal' }
      }

      macd_chart.setOption('series', macd_series_labels);
      macd_chart = macd_chart.build();
      sheet.insertChart(macd_chart);
      SpreadsheetApp.flush();
      var macd_chart_image = macd_chart.getAs('image/png');
      var inserted_image = doc_body.insertImage(1,macd_chart_image);

      spreadsheetUrl = ss.getUrl();
      //doc_body.appendParagraph('Here is the link to the Google Spreadsheet: ' + spreadsheetUrl);

      inserted_image.setLinkUrl(spreadsheetUrl); //put link to spreadsheet in image
    }

  } else if (visualization == "Table")
  {

    var range = sheet.getRange(2, 1, 1, 2).getDataRegion(SpreadsheetApp.Dimension.ROWS); //start at A1, 1 row, 2 columns
    var values = range.getValues();
    var backgroundColors = range.getBackgrounds();
    var styles = range.getTextStyles();

    // Position to paste data in Google Docs
    var doc_body = doc.getBody();
    for (var i = 0; i < values.length; i++) {
      // If the value is a Date object, format it as a string
      if (values[i][0] instanceof Date) {
        // Format the date in column A as a string so that we can turn it into table in the document
        values[i][0] = Utilities.formatDate(values[i][0], Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
    }
    var table = doc_body.insertTable(1,values);
    //table.setBorderWidth(0);
    for (var i = 0; i < table.getNumRows(); i++) {
      const tableRow = table.getRow(i);
      for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
        const cell = tableRow.getCell(j);
        //allow no cell padding
        cell.setPaddingTop(0);
        cell.setPaddingBottom(0);
        cell.setPaddingLeft(0);
        cell.setPaddingRight(0);
      }
    }

    spreadsheetUrl = ss.getUrl();
    doc_body.insertParagraph(1,'Here is the link to the Google Spreadsheet: ' + spreadsheetUrl);
  }

  //now calculate statistics and put them in sheets and doc
  if ((statistics >>> 6) % 2 == 1)
  {
    calcCurrentPrice(sheet,stockTicker,doc,21)
  }
  if ((statistics >>> 5) % 2 == 1)
  {
    calcEndPrice(sheet,doc,20)
  }
  if ((statistics >>> 4) % 2 == 1)
  {
    calcStartPrice(sheet,doc,19);
  }
  if ((statistics >>> 3) % 2 == 1)
  {
    calcRelativePrice(sheet,doc,18);
  }
  if ((statistics >>> 2) % 2 == 1)
  {
    calcAvg(sheet,doc,17);
  }
  if ((statistics >>> 1) % 2 == 1)
  {
    calcMin(sheet,doc,16);
  }
  if (statistics % 2 == 1)
  {
    calcMax(sheet,doc,15);
  }


  
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