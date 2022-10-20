
function copyPnL(){
  var uri = "https://docs.google.com/spreadsheets/d/1pNObBNUMu64yYJ2rWBBggWQE0nKnvNNmCel45rE6f2g/edit#gid=146678684";
  var sheetname = "Hum Budget";
  var range = "A1:BS500";
  var offsetx = 1;
  var offsety = 1;
  COPYSHEET(uri, sheetname, range, offsetx, offsety);
}

function COPYSHEET(spreadsheet_uri, sheetname, range, offsetx = null, offsety = null){

  if (offsetx == null){ offsetx = 1; }
  if (offsety == null){ offsety = 1; }

  var file = SpreadsheetApp.openByUrl(spreadsheet_uri);
  var sheet = file.getSheetByName(sheetname);
  var range = sheet.getRange(range);
  var values = range.getValues();
  
  SpreadsheetApp.getActive().getActiveSheet().getRange(offsetx, offsety,values.length, values[0].length).clear();
  SpreadsheetApp.getActive().getActiveSheet().getRange(offsetx, offsety,values.length, values[0].length).setValues(values);
}

/**
 * Paste this code in the editor. Run a first time via the menu "COPY > Update sheets", it gives you a configuration sheet, configure, then run again to copy sheets
 */
function updatesheets(){
  var confsheet = SpreadsheetApp.getActive().getSheetByName("COPY Conf");
  // create configuration page if not exists
  if (!confsheet){
    confsheet = SpreadsheetApp.getActive().insertSheet();
    confsheet.setName("COPY Conf");
    confsheet.getRange(1,1).setValue("COPY SCRIPT CONFIGURATION").setFontSize(16).setFontWeight("bold");
    confsheet.getRange(2,1,1,8).setValues([["Update", "URL", "Sheet Name", "Range", "Destination Sheet", "Offset Row", "Offset Column", "Last Update"]]).setFontWeight("bold");
    confsheet.setColumnWidths(1,1, 100);
    confsheet.setColumnWidths(2,2, 400);
    confsheet.setColumnWidths(3,5, 200);
    confsheet.setColumnWidths(6,7, 100);

    confsheet.getRange(3,1,13,1).insertCheckboxes(); 
  }

  // loop on configuration lines
  var conf = confsheet.getRange(3,1,13,8).getValues();
  conf = conf.map((confobj) => {
    if (!confobj[0]) return confobj; // checkbox uncheck

    offsetx = confobj[5];
    offsety = confobj[6];
    if (offsetx == null){ offsetx = 1; }
    if (offsety == null){ offsety = 1; }

    var file = SpreadsheetApp.openByUrl(confobj[1]);
    var sheet = file.getSheetByName(confobj[2]);
    var range = sheet.getRange(confobj[3]);
    var values = range.getValues();

    var destsheet = SpreadsheetApp.getActive().getSheetByName(confobj[4]); 
    if (!destsheet){
      destsheet = SpreadsheetApp.getActive().insertSheet();
      destsheet.setName(confobj[4]);
    }

    // clear destination sheet
    destsheet.getRange(offsetx, offsety,values.length, values[0].length).clear();
    destsheet.getRange(offsetx, offsety,values.length, values[0].length).setValues(values);

    confobj[7] = "Maj : "+Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
    return confobj;
  });

  // update "last update" field
  confsheet.getRange(3,1,13,8).setValues(conf);
}


function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu("COPY")
  .addItem("Update sheets", "updatesheets")
  .addToUi();
}
