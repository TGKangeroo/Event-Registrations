
//Add Triggers --------------------------------------------------------------------------------------------------------------------------------------------------------//
function resetTriggers() {
    removeTriggers();
    ScriptApp.newTrigger('onESNEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
    ScriptApp.newTrigger('onESNSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();
}
//Remove Triggers --------------------------------------------------------------------------------------------------------------------------------------------------------//
function removeTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    }
}

//Print Triggers --------------------------------------------------------------------------------------------------------------------------------------------------------//
function printTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        Logger.info(allTriggers[i]);
    }
}
//on Sheet Edit Trigger --------------------------------------------------------------------------------------------------------------------------------------------------------//
function onESNEdit(e) {
    var editedSheet = e.source.getActiveSheet();
    if( editedSheet.getName() != registerSheet.getName() ) {
        return;
    }
    var range = e.range

    if(range.getNumColumns() != 1 || range.getNumRows() != 1) {
        showAlert('Warning', 'Only change one row at a time!');
        return;
    }
    var rowId = range.getRow(); // row used for inserting into the google sheet

    var row = registerSheet.getRange(rowId, 1, 1, registerSheet.getLastColumn());
    var paidCell = row.getCell(1, indexOfPaid);
    var whenCell = row.getCell(1, indexOfPaidWhen);
    var paidLoc = row.getCell(1,indexOfPaidLoc);

    
    var event_max_participants = options["MAX_PARTICIPANTS"];
    var amount_total_part = updateTotalParticipants();
    //check if the changed value is on the paid row and if it's changed to yes

    if (range.getColumn() == indexOfPaid && indexOfPaid != -1) {
        switch (e.value) {
            case "yes":
                if(options["PAID_EVENT"] && 
                    (whenCell.getValue() == "" || paidLoc.getValue() == "")) {
                        
                    paidCell.setValue('no');
                    paidCell.setComment('you have to fill When Where column first');
                    return;
                }
                if (amount_total_part > event_max_participants) {
                    paidCell.setValue('no');
                    paidCell.setComment('Denied max participants reached');
                    row.getCell(1, getColumnId("last Edited")).setValue(new Date());
                    SpreadsheetApp.flush();
                    return;
                }
                row.setBackground("MediumSeaGreen");
                if (options["AUTO_CONF_MAIL"] == true) {
                    sendconfirmationEmail(row);
                }
                
                if (amount_total_part == event_max_participants) {
                    paidCell.setComment('Warning, last participant!');
                    if (options["CLOSE_WHEN_FULL"]) {
                        closeForm();
                    }
                }
                break;
            case "no":
                row.setBackground("white");
                sendRegisterEmail(row); // try to resend mail if not sent
                break;
            case "cancelled":
                row.setBackground("red");
                break;
            case "refunded":
                row.setBackground("lightBlue");
                break;
        }
        row.getCell(1, getColumnId("last Edited")).setValue(new Date());
        totalPriceToBePaid(row);
        updatePrices();
        SpreadsheetApp.flush();
    }
}

//on Form Submit Trigger --------------------------------------------------------------------------------------------------------------------------------------------------------//
function onESNSubmit(e) {
    if (indexOfPaid == -1) {
        makePayAndEditedRow();
    }
    console.info('onSubmit running as', getUserEmail()); //DEBUG
    var range = e.range;
    var rowId = range.getRow(); // row used for inserting into the google sheet

    var row = registerSheet.getRange(rowId, 1, 1, registerSheet.getLastColumn());
    Logger.info("row" + row);
    var cell = row.getCell(1, indexOfPaid);

    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['yes', 'no', 'cancelled', 'refunded','Waitinglist'], false).build();
    cell.setDataValidation(rule);
    
    var paidLoc = row.getCell(1,indexOfPaidLoc);
    paidLoc.setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(
            options["PAYMENT_LOCATION"].split(","), false).build());
    

    var event_max_participants = options["MAX_PARTICIPANTS"];
    var amount_total_part = 0;
    amount_total_part = countParticipants() + 1; //add +1 cause entry is still not persisted
    if (options["PAID_EVENT"] == true) {
        if (cell.getValue() != "cancelled" && cell.getValue() != "refunded") {
            cell.setValue('no');
        }
        if (options["AUTO_REG_MAIL"] == true) {
            sendRegisterEmail(row);
        }
    } else {
        cell.setValue('yes');
        updatePrices();
        if (options["AUTO_CONF_MAIL"] == true) {
            sendconfirmationEmail(row);
        }
        if ((amount_total_part == event_max_participants ||
            amount_total_part > event_max_participants) &&
            event_max_participants != "0" && options["CLOSE_WHEN_FULL"]) {
            Logger.log("event reached max participants");
            closeForm();
        }
    }

    totalPriceToBePaid(row);
    SpreadsheetApp.flush();
}

