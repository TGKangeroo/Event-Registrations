function cleanList() {
    var start, end;
    start = 2;
    end = printSheet.getLastRow();//Number of last row with content
    //blank rows after last row with content will not be deleted
    if (end > 1) {
        printSheet.deleteRows(start, end);
    }
}
//Remake the print list with all participants --------------------------------------------------------------------------------------------------------------------------------------------------------//
function refreshPrintList() {

    cleanList();
    if(registerSheet == null) {
        // form not created yet
        return;
    }
    var printSheetRange = printSheet.getRange(1, 1, 1, printSheet.getLastColumn());
    var printSheetColumns = printSheetRange.getValues();

    var registrations = registerSheet.getDataRange().getValues();
    var lastRow = registerSheet.getLastRow();
    var lastCol = printSheet.getLastColumn();
    for (var y = 1; y < lastRow; ++y) {
        var printrow = [];
        if (getByNameData(registrations, "Paid", y) == "yes") {
            for (var i = 0; i < lastCol; i++) {
                printrow[i] = getByNameData(registrations, printSheetColumns[0][i], y);
            }
            printSheet.appendRow(printrow);
        }
    }
}
