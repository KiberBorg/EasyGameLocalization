function onInstall(e) {
  onOpen(e);
}

/**
 *
 * Toolbar menu creation.
 *
 * Called on worbook opening.
 *
 **/
function onOpen() {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Start translating', 'showSidebar')
        .addItem('About', 'showAbout')
        .addToUi();
}

/**
 *
 * Sidebar title, content & size.
 *
 **/
function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('index')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Easy Game Localization')
        .setWidth(300);

    // Open sidebar
    SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  var html = HtmlService.createHtmlOutputFromFile('about')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('About')
      .setWidth(250)
      .setHeight(500);
  SpreadsheetApp.getActive().show(html);
}

function startTranslation(overwrite) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Translation in progress...", "", -1);
    try {
      var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      translate(activeSpreadsheet, overwrite);
    } catch (err) {
      SpreadsheetApp.getActiveSpreadsheet().toast("An error occured:" + err);
    }
}
function translate(activeSpreadsheet, overwrite)
{
  var lrow = activeSpreadsheet.getLastRow();
  var lcol = activeSpreadsheet.getLastColumn();
  var sourcelanguage = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 2).getValue();
  var sourcelanguagewords = [];
  for (var i = 2; i <= lrow; i++) 
  {
    if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, 2).getValue().length > 1) {
      var activeCellText = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, 2).getValue();
      sourcelanguagewords.push(activeCellText);
    }
  }
  for (var j = 3; j <= lcol; j++) {
    var targetlanguage = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, j).getValue();
    for (var i = 2; i <= lrow; i++) {
      if(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, j).getValue() == "" || overwrite)
      {
        var sourceword = sourcelanguagewords[i-2];
        var activeCellTranslation = LanguageApp.translate(sourceword, sourcelanguage, targetlanguage);
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, j).setValue(activeCellTranslation);
        activeSpreadsheet.toast("Translated to " + targetlanguage, "", 2);
      }
    }
  }
  activeSpreadsheet.toast("Done translating.", "Sucess", 4);
}
