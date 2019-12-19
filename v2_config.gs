
function readConfig() {
    var options = {};
    var optionSheet = SpreadsheetApp.getActiveSpreadsheet();
    var values = ss.getRange("Options!A3:C50").getValues();
    
    for (var row in values) {

        var name = values[row][0];
        if(name != "") {
            options[name] = values[row][2];
        }
    }
    return options;
    
}
