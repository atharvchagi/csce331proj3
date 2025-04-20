//////////////////////
//indicator functions
//////////////////////
function calcMovingAverage(sheet,column) {
  //20 day moving average
  var last_column = column; //next empty column
  sheet.getRange(1, last_column).setValue("Moving Average");
  sheet.getRange(21, last_column).setFormula("=AVERAGE(B2:B21)");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(21,last_column, last_row-20);
  sheet.getRange(21, last_column).copyTo(fillDownRange);
}

function calcRSI(sheet,starting_column) {
  //14 day RSI
  var last_column = starting_column;
  sheet.getRange(1, last_column).setValue("Change");
  sheet.getRange(3, last_column).setFormula("=B3-B2");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(3,last_column, last_row-2);
  sheet.getRange(3, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("Gain");
  sheet.getRange(3, last_column).setFormula("=IF(D3>0,D3,0)");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(3,last_column, last_row-2);
  sheet.getRange(3, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("Loss");
  sheet.getRange(3, last_column).setFormula("=IF(D3<0,-D3,0)");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(3,last_column, last_row-2);
  sheet.getRange(3, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("Average Gain");
  sheet.getRange(16, last_column).setFormula("=AVERAGE(E3:E16)");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(16,last_column, last_row-15);
  sheet.getRange(16, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("Average Loss");
  sheet.getRange(16, last_column).setFormula("=AVERAGE(F3:F16)");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(16,last_column, last_row-15);
  sheet.getRange(16, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("RS");
  sheet.getRange(16, last_column).setFormula("=G16/H16");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(16,last_column, last_row-15);
  sheet.getRange(16, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("RSI");
  sheet.getRange(16, last_column).setFormula("=IF(H16=0, 100, 100-(100/(1+I16)))");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(16,last_column, last_row-15);
  sheet.getRange(16, last_column).copyTo(fillDownRange);
}

function calcMACDandSignal(sheet,starting_column) {
  //MACD (12,26) and Signal (9)
    var last_column = starting_column;
  sheet.getRange(1, last_column).setValue("12 Day EMA");
  sheet.getRange(2, last_column).setValue("=2/(12+1)"); //calculate multiplier and put in K2 cell
  sheet.getRange(13, last_column).setFormula("=AVERAGE(B2:B13)");
  sheet.getRange(14, last_column).setFormula("=(B14-K13)*$K$2+K13");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(14,last_column, last_row-13);
  sheet.getRange(14, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("26 Day EMA");
  sheet.getRange(2, last_column).setFormula("=2/(26+1)");
  sheet.getRange(27, last_column).setFormula("=AVERAGE(B2:B27)");
  sheet.getRange(28, last_column).setFormula("=(B28-L27)*$L$2+L27");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(28,last_column, last_row-27);
  sheet.getRange(28, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("MACD"); //MACD = 12 Day EMA of stock close price - 26 Day EMA of stock close price
  sheet.getRange(27, last_column).setFormula("=K27-L27");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(27,last_column, last_row-26);
  sheet.getRange(27, last_column).copyTo(fillDownRange);

  var last_column = last_column + 1;
  sheet.getRange(1, last_column).setValue("Signal"); //9 EMA of MACD is signal
  var multiplier = 0.2; //multiplier = 2/(9+1) = 0.2. Will show up in chart, so can't put it in spreadsheet
  sheet.getRange(35, last_column).setFormula("=AVERAGE(M27:M35)");
  sheet.getRange(36, last_column).setFormula("=(M36-N35)*"+multiplier+"+N35");
  var last_row = sheet.getLastRow();
  var fillDownRange = sheet.getRange(36,last_column, last_row-35);
  sheet.getRange(36, last_column).copyTo(fillDownRange);
}

//////////////////////
//statistic functions
//////////////////////
function calcMax(sheet, doc, column) {
  //max
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Max");
  sheet.getRange(2, column).setFormula("=MAX(B:B)");
  doc_body.insertParagraph(1, "Maximum: " + sheet.getRange(2,column).getValue());
}

function calcMin(sheet, doc, column) {
  //min
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Min");
  sheet.getRange(2, column).setFormula("=MIN(B:B)");
  doc_body.insertParagraph(1, "Minimum: " + sheet.getRange(2,column).getValue());
}

function calcAvg(sheet, doc, column) {
  //average
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Avg");
  sheet.getRange(2, column).setFormula("=AVERAGE(B:B)");
  doc_body.insertParagraph(1, "Average: " + sheet.getRange(2,column).getValue());
}

function calcRelativePrice(sheet, doc, column) {
  //relative change in price/attribute
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Relative Change");
  var last_row = sheet.getLastRow();
  sheet.getRange(2, column).setFormula("B" + last_row + "/" + "B2");
  doc_body.insertParagraph(1, "Relative Price: " + sheet.getRange(2,column).getValue());
}

function calcStartPrice(sheet, doc, column) {
  //start price/attribute
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Start");
  sheet.getRange(2, column).setFormula("B2");
  doc_body.insertParagraph(1, "Start: " + sheet.getRange(2,column).getValue());
}

function calcEndPrice(sheet, doc, column) {
  //end price/attribute
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("End");
  var last_row = sheet.getLastRow();
  sheet.getRange(2, column).setFormula("B" + last_row);
  doc_body.insertParagraph(1, "End Price: " + sheet.getRange(2,column).getValue());
}

function calcCurrentPrice(sheet, stockTicker, doc, column) {
  //current price/attribute
  var doc_body = doc.getBody();
  sheet.getRange(1, column).setValue("Current");
  sheet.getRange(2, column).setFormula("GOOGLEFINANCE(" + stockTicker + ","+  '"price")' );
  doc_body.insertParagraph(1, "Current Price: " + sheet.getRange(2,column).getValue());
}
