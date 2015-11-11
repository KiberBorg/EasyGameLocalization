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
        .setTitle('Easy Localization')
        .setWidth(300);

    // Open sidebar
    SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  var html = HtmlService.createHtmlOutputFromFile('about')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('About')
      .setWidth(250)
      .setHeight(450);
  SpreadsheetApp.getActive().show(html);
}
function showFormatting()
{
  var html = HtmlService.createHtmlOutputFromFile('formatinfo')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('How to use')
      .setWidth(400)
      .setHeight(420);
  SpreadsheetApp.getActive().show(html);
}
function translate()
{
  SpreadsheetApp.getActiveSpreadsheet().toast("Translation in progress...", "", 3);
  var lrow = SpreadsheetApp.getActiveSpreadsheet().getLastRow();
  var lcol = SpreadsheetApp.getActiveSpreadsheet().getLastColumn();
  var sourcelanguage = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 2).getValue();
  var sourcelanguagewords = [];
  var errorcounter = 0;
  var sourceIncorrect = false;
   var ui = SpreadsheetApp.getUi();
  for (var i = 2; i <= lrow; i++) 
  {
    if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, 2).getValue().length > 1) {
      var activeCellText = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, 2).getValue();
      sourcelanguagewords.push(activeCellText);
    }
  }
 if(sourcelanguagewords.length < 1)
 {
   ui.alert("Your sheet has no words to translate. Please look at the 'How to use' section if you're unsure what format your sheet has to be in.");
   return;
 }
   try {
      var sourcelanguageTrial = LanguageApp.translate('Hello', sourcelanguage, 'es');
    } catch (err) {
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 2).setBackground("#FF4A4A");
      sourceIncorrect = true;
    }
  if(!sourceIncorrect)
  {
  for (var j = 3; j <= lcol; j++) {
    var targetlanguage = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, j).getValue();
    for (var i = 2; i <= lrow; i++) {
      if(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, j).getValue() == "")
      {
        var sourceword = sourcelanguagewords[i-2];
        try {
       var activeCellTranslation = LanguageApp.translate(sourceword, sourcelanguage, targetlanguage);
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i, j).setValue(activeCellTranslation);
        SpreadsheetApp.getActiveSpreadsheet().toast("Translated to " + targetlanguage, "", 2);
    } catch (err) {
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, j).setBackground("#FF4A4A");
      errorcounter++;
      break;
      }
    }
    }
  }
  }
  if(errorcounter > 0)
  {
  ui.alert(
    errorcounter+' of the columns had non identifiable language codes and could not be translated.'+
    ' They have been marked red so you know where the errors are',
    ui.ButtonSet.OK);
  }else if(sourceIncorrect)
  {
    ui.alert(
    'Your main language seems not to be set-up correctly, please provide a valid language code in order for this to work.'+
    ' The cell has been marked red.',
    ui.ButtonSet.OK);
  }
  else{
  SpreadsheetApp.getActiveSpreadsheet().toast("Done translating.", "Sucess", 4);
  }
}

function createExample()
{
   try {
  var newsheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Example Localization Sheet");
  var languageCodes = ["Identifiers / Language Codes", "en", "de", "fr", "es", "ru"];
  var identifiers = ["greeting", "short_sentence", "arthur_presentation"];
  var sourceWords = ["Hello", "A long train.", "Hello my name is Arthur"];
  for(var i = 0; i< languageCodes.length; i++)
  {
    newsheet.getRange(1, i+1).setValue(languageCodes[i]);
    newsheet.getRange(1, i+1).setFontWeight('bold');
  }
  for(var i = 0; i< identifiers.length; i++)
  {
    newsheet.getRange(i+2, 1).setValue(identifiers[i]);
  }
   for(var i = 0; i< sourceWords.length; i++)
  {
    newsheet.getRange(i+2, 2).setValue(sourceWords[i]);
  }
     SpreadsheetApp.getActiveSpreadsheet().toast("Added Localization Sheet.", "Sucess", 3);
   }catch (err) {
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        'There already is an example sheet.',
        ui.ButtonSet.OK);
    }
}