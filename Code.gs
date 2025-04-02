//https://developers.google.com/apps-script/guides/dialogs
//function onOpen() {
//  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
//    .createMenu("Custom Menu")
//    .addItem("Show alert", "showAlert")
//    .addToUi();
//}

function onOpen() {
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('My custom sidebar');
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
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

function showStockAnalyzer() {
    var html = HtmlService.createHtmlOutputFromFile('StockAnalyzerUI')
        .setWidth(400)
        .setHeight(300);
    DocumentApp.getUi().showModalDialog(html, 'Stock Analyzer');
}