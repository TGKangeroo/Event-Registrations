//variables for sheet declaration --------------------------------------------------------------------------------------------------------------------------------------------------------//
var ss = SpreadsheetApp.getActiveSpreadsheet();
var optionSheet = ss.getSheetByName("Options");
var priceSheet = ss.getSheetByName("Prices");
var printSheet = ss.getSheetByName("Print list");
var budgetSheet = ss.getSheetByName("Budget");
var registerSheet = ss.getSheetByName("Registrations");


var registerHeaders = []
if (registerSheet != null) {
    registerHeaders = registerSheet.getRange("A1:AA1").getValues()[0];
}

var _drafts = []; // Lazy load draft mails
function getDraftMessages() {
    if (_drafts.length == 0) {
        _drafts = GmailApp.getDraftMessages()
    }
    return _drafts;
}
var _prices = []; // lazy load
//returns array of all prices written in the yellow optionsheet block --------------------------------------------------------------------------------------------------------------------------------------------------------//
function getAllPrices() {
    if (_prices.length == 0) {
        _prices = priceSheet.getRange("A3:E19").getValues();
        _prices = _prices.filter( function(price) {
            return price[0] != "";
        });
    }
    return _prices;
}
//adds the price a participant has to pay to the registration sheet --------------------------------------------------------------------------------------------------------------------------------------------------------//
function totalPriceToBePaid(row) {
    var price = calculatePrice(row);
    row.getCell(1, indexOfTotal).setValue(price);
}
function protectFirstRow() {
    var range = registerSheet.getRange(1,1,1,registerSheet.getLastColumn())
    var protection = range.protect().setDescription('Script protected!');
    protection.removeEditors(protection.getEditors());
    protection.setDomainEdit(false);
    protection.setWarningOnly(true);
    SpreadsheetApp.flush();
}

//Add last edited and paid columns to registration sheet --------------------------------------------------------------------------------------------------------------------------------------------------------//
function makePayAndEditedRow() {
    var idxLast = registerSheet.getLastColumn()

    registerSheet.getRange(1, idxLast + 1).setValue('to be paid');
    registerSheet.getRange(1, idxLast + 2).setValue('Paid');
    registerSheet.getRange(1, idxLast + 3).setValue('When');
    registerSheet.getRange(1, idxLast + 4).setValue('Where');
    registerSheet.getRange(1, idxLast + 5).setValue('How Much');
    registerSheet.getRange(1, idxLast + 6).setValue('Comment');
    registerSheet.getRange(1, idxLast + 7).setValue('last Edited');
    registerSheet.getRange(1, idxLast + 8).setValue('Script');

    protectFirstRow();


    SpreadsheetApp.flush();
    registerHeaders = registerSheet.getDataRange().getValues()[0];

    indexOfPaid = getColumnId("Paid");
    indexOfPaidWhen = getColumnId("When");
    indexOfPaidLoc = getColumnId("Where");
    indexOfTotal = getColumnId("to be paid");
    indexOfScript = getColumnId("Script");
}
//retrieves the number of a column based on the column name --------------------------------------------------------------------------------------------------------------------------------------------------------//
function getColumnId(colName) {
    var col = registerHeaders.indexOf(colName);
    if (col != -1) {
        return col + 1;
    } else {
        return -1;
    }
}

var indexOfPaid = getColumnId("Paid");
var indexOfPaidWhen = getColumnId("When");
var indexOfPaidLoc = getColumnId("Where");
var indexOfTotal = getColumnId("to be paid");
var indexOfScript = getColumnId("Script");



//searches a registration value based on column name and row --------------------------------------------------------------------------------------------------------------------------------------------------------//
function getByNameRow(colName, row) {
    var colId = getColumnId(colName);
    if (colId < 0) {
        Logger.log("could not find colId for " + colName);
        return "not found:" + colName;
    }
    return row.getCell(1, colId).getValues()[0];
}
function getByName(colName, rowId) {
    var data = registerSheet.getDataRange().getValues();
    return getByNameData(data, colName, rowId);
}
function getByNameData(data, colName, row) {
    var col = data[0].indexOf(colName);
    if (col != -1) {
        if (data[row] != null) {
            if (data[row][col] != null) {
                return data[row][col];
            } else {
                return "";
            }
        }
    }
}
//updates the amount field 30B --------------------------------------------------------------------------------------------------------------------------------------------------------//
function updatePrices() {
    var idxAmount = 4;
    var idxField = 1;
    var idxValue = 2;
    var prices = getAllPrices();
    var amounts = new Array(prices.length);
    for (var i = 0; i < prices.length; i++) {
        amounts[i] = new Array(1);
        amounts[i][0] = 0;
    }
    var dataRange = registerSheet.getRange(2, 1, registerSheet.getLastRow() - 1, registerSheet.getLastColumn()); // let it read more columns than are being used, it might mess up otherwise
    var registrations = dataRange.getValues();
    // Fetch values for each row in the Range.
    var paid = getColumnId("Paid") - 1; //array idx

    for (var i = 0; i < registrations.length; ++i) {
        if (registrations[i][paid] == "yes") {
            for (var y = 0; y < prices.length; y++) {
                var index = registerHeaders.indexOf(prices[y][idxField]);
                if (registrations[i][index] == prices[y][idxValue]) {
                    amounts[y][0] = amounts[y][0] + 1;
                }
            }
        }
    }
    priceSheet.getRange(3, idxAmount + 1, prices.length).setValues(amounts);
}
//Calculate the price per participant --------------------------------------------------------------------------------------------------------------------------------------------------------//
function calculatePrice(row) {
    var idxField = 1;
    var idxValue = 2;
    var idxPrice = 3;

    var prices = getAllPrices();
    var pay = 0;
    for (var i = 0; i < prices.length; i++) {
        if (!(isNaN(prices[i][idxPrice])) && prices[i][idxPrice] != "") {
            if (prices[i][3] == "Base Price") {
                pay = pay + prices[i][idxPrice];
            } else {
                if (getByNameRow(prices[i][idxField], row) == prices[i][idxValue]) {
                    pay = pay + prices[i][idxPrice];
                }
            }
        }
    }
    return pay;
}

//Calculate the price for a paypal user + generates payment link--------------------------------------------//

function calculatePayPal(row){
  var price = calculatePrice(row);
var paypalUrl=PAYPAL_URL;
  var email = 'finance@esnuniwien.com';
  var description = options["EVENT_TITLE"];
 
  description = description + " " + getByNameRow('First name', row) + " " + getByNameRow('Last name', row);


  var payPalExtraCost = 1.034;

  price = price * payPalExtraCost + 0.35;

  description = encodeURIComponent(description);

  paypalUrl = paypalUrl.replace("EMAIL",email);
  paypalUrl = paypalUrl.replace("PRICE",price);
  paypalUrl = paypalUrl.replace("EVENT",description);

  return  paypalUrl;
}
//check if the registration end date equals to todays date --------------------------------------------------------------------------------------------------------------------------------------------------------//
function checkEndDate() {
    var enddate = getFieldValue('script_registration_close_date');
    if (enddate != null) {
        enddate = Utilities.formatDate(new Date(enddate), "Europe/Vienna", "dd-yyyy-MM");
        var today = Utilities.formatDate(new Date(), "Europe/Vienna", "dd-yyyy-MM");
        if (today == enddate || today > enddate) {
            closeForm();
            removeTriggers();
            makeTriggers();
        }
    }
}
//Count the amount of paid participants --------------------------------------------------------------------------------------------------------------------------------------------------------//
function countParticipants() {
    var dataRange = registerSheet.getRange(2, indexOfPaid, registerSheet.getLastRow(), indexOfPaid); // let it read more columns than are being used, it might mess up otherwise
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    var counter = 0;
    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        if (row[0] == "yes") {
            counter++;
        }
    }
    return counter;
}
function updateTotalParticipants() {
    var participants = countParticipants();
    priceSheet.getRange("C20").setValue(participants);
    return participants;
}

//test function for development --------------------------------------------------------------------------------------------------------------------------------------------------------//
function showAlert(title, msg) {
    var htmlOutput = HtmlService
        .createHtmlOutput(msg)
        .setWidth(250)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);

    Logger.log(title + msg);
}

