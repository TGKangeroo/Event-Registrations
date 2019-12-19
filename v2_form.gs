var options = readConfig();

//Stops accepting responses to the google form --------------------------------------------------------------------------------------------------------------------------------------------------------//
function closeForm() {
    var form = FormApp.openByUrl(optionSheet.getRange('G22').getValue());
    form.setAcceptingResponses(false);
    form.setCustomClosedFormMessage("I'm sorry but we're currently not taking anymore registrations. This means that either we've reached the maximum amount of participants, or the registration deadline has passed.");
}
//Opens the google form for responses --------------------------------------------------------------------------------------------------------------------------------------------------------//
function openForm() {
     var form = FormApp.openByUrl(optionSheet.getRange('G22').getValue());
    form.setAcceptingResponses(true);
}
function isFormCreated() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
        var name = sheets[i].getName()
        if ( name == "Registrations" || name.indexOf("responses")>-1 ) {
            Logger.log(name);
            return true;
        }
    }
    return false;
}

function renameResponsesSheet() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
        var name = sheets[i].getName()
        if ( name.indexOf("responses")>-1 ) {
            sheets[i].setName("Registrations");
        }
    }
}
//Make google form ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
function makeForm() {
    if (isFormCreated()) {
        Logger.log("form already created")
        resetTriggers();
        createDraftMails();
        return;
    } else {
        Logger.log("creating form");
    }

    form = FormApp.create(options["EVENT_TITLE"])
        .setDescription(options["EVENT_DESCRIPTION"])
        .setConfirmationMessage('Thanks for registering!')
        .setAllowResponseEdits(false)
        .setAcceptingResponses(true)
        .setRequireLogin(false);
    
  

    var questions = getAllQuestions();
    for (var i = 0; i < questions.length; i++) {
        var question = questions[i];
        
        var required = question.required;;
        var title = question.title;
        var desc = replaceConfVars(question.desc);
        
        qoptions = question.options.map(function (v) { return replaceConfVars(v); });
 
        switch(question.type) {
            case "text": makeTextItem(form, title, desc, required); break;
            case "email": makeEmailItem(form, title, desc, required); break;
            case "dropdown": makeDropdownItem(form, title, desc, qoptions, required); break;
            case "radiobutton": makeRadioButtonItem(form, title, desc, qoptions, required); break;
            case "checkbox": makeCheckBoxItem(form, title, desc, qoptions, required); break;
            case "date": makeDateItem(form, title, desc, required); break;
            case "time": makeTimeItem(form, title, desc, required); break;
            case "datetime": makeDateTimeItem(form, title, desc, required); break;
            case "duration": makeDurationItem(form, title, desc, required); break;
        }
    }

    SpreadsheetApp.flush();

    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  ss.addEditor("jenske130@gmail.com");
    renameResponsesSheet();
    resetTriggers();
    
    //Sheet and form sharing permissions 
     var thisSheet = DriveApp.getFileById(ss.getId());
     thisSheet.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT)
    //sharingIsCaring(Form);
    
    optionSheet.getRange("G21").setValue(form.getPublishedUrl());
  optionSheet.getRange("G22").setValue(form.getEditUrl());
    optionSheet.getRange("G23").setValue(ss.getUrl());
    SpreadsheetApp.flush();

    createDraftMails();

}
function replaceConfVars(string) {
    for(var key in  options) {
        string = replaceTemplateKey(string, key, options[key])
    }
    return string;
}
function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // $& means the whole matched string
}

function replaceTemplateKey(msg,key,value) {
    return msg.replace(new RegExp(escapeRegExp("["+key+"]"), 'g'), value);
}
//Opens the google form for responses --------------------------------------------------------------------------------------------------------------------------------------------------------//
function addCancellationPolicy(form) {
    var cancellation_choices = "I hereby agree to the cancellation policy";
    //var choices="I hereby agree that my data will be processed by ESN Vienna and its members in accordance with <a href='" +  dsgvo_link + "'>the data protection rules of ESN Austria</a>";
    var cancellation_description = "Cancellation with 100% refund is possible up to 14 days before the beginning of the event. It is allowed to pass a paid spot under the same conditions to another person up to 24h before the beginning of the event via emailing the organizers.";
    makeRadioButtonItem(form, "Cancellation Policy", cancellation_description, cancellation_choices.split(','), true)
    script_form_fields_amount = optionSheet.getRange('B30').getValue();
    optionSheet.getRange(script_form_fields_amount + 4, 10).setValue("Cancellation Policy");
    optionSheet.getRange(script_form_fields_amount + 4, 11).setValue('radiobutton');
    optionSheet.getRange(script_form_fields_amount + 4, 13).setValue('TRUE');
    optionSheet.getRange(script_form_fields_amount + 4, 12).setValue(cancellation_choices);
    optionSheet.getRange(script_form_fields_amount + 4, 14).setValue(cancellation_description);
}
//Make form question with email input field -----------------------------------------------------
function makeEmailItem(form, title, Description, required, item) {
    if (item == null) {
        item = form.addTextItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
    var emailValidation = FormApp.createTextValidation()
        .requireTextIsEmail()
        .build();
    item.setValidation(emailValidation)
    var emailValidation = FormApp.createTextValidation()
        .requireTextIsEmail()
        .build();
    item.setValidation(emailValidation)
}
//Make form question with text input field -----------------------------------------------------
function makeTextItem(form, title, Description, required, item) {
    if (item == null) {
        item = form.addTextItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
}
//Make form question with dropdown field -----------------------------------------------------
function makeDropdownItem(form, title, Description, choices, required, item) {
    var process = "update";
    if (item == null) {
        var item = form.addListItem();
        process = "insert";
    }
    if (process == "update") {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.asListItem().createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.asListItem().setTitle(title)
            .setChoices(arrayOfItems)
            .setHelpText(Description);
    } else {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.setTitle(title)
            .setChoices(arrayOfItems)
            .setHelpText(Description)
            .setRequired(required);
    }
}
//Make form question with checkbox field -----------------------------------------------------
function makeCheckBoxItem(form, title, Description, choices, required, item) {
    process = "update";
    if (item == null) {
        var item = form.addCheckboxItem();
        process = "insert";
    }
    if (process == "update") {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.asCheckboxItem().createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.asCheckboxItem().setTitle(title)
            .setChoices(arrayOfItems)
            .showOtherOption(false)
            .setHelpText(Description)
            .setRequired(required);
    }
    else {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.setTitle(title)
            .setChoices(arrayOfItems)
            .showOtherOption(false)
            .setHelpText(Description)
            .setRequired(required);
    }
}
//Make form question with radio button field -----------------------------------------------------
function makeRadioButtonItem(form, title, Description, choices, required, item) {
    process = "update";
    if (item == null) {
        var item = form.addMultipleChoiceItem();
        process = "insert";
    }
    if (process == "update") {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.asMultipleChoiceItem().createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.asMultipleChoiceItem().setTitle(title)
            .setChoices(arrayOfItems)
            .setHelpText(Description)
            .showOtherOption(false)
            .setRequired(required);
    } else {
        var arrayOfItems = [];
        for (var i = 0; i < choices.length; i++) {
            var thisValue = choices[i];
            var newItem = item.createChoice(thisValue);
            arrayOfItems.push(newItem);
        }
        item.setTitle(title)
            .setChoices(arrayOfItems)
            .setHelpText(Description)
            .showOtherOption(false)
            .setRequired(required);
    }
}
//Make form question with date input field -----------------------------------------------------
function makeDateItem(form, title, Description, required, item) {
    if (item == null) {
        var item = form.addDateItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
}
//Make form question with date item input field -----------------------------------------------------
function makeDateTimeItem(form, title, Description, required, item) {
    if (item == null) {
        var item = form.addDateTimeItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
}
//Make form question with time input field -----------------------------------------------------
function makeTimeItem(form, title, Description, required, item) {
    if (item == null) {
        var item = form.addTimeItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
}
//Make form question with duration input field -----------------------------------------------------
function makeDurationItem(form, title, Description, required, item) {
    if (item == null) {
        var item = form.addDurationItem();
        item.setRequired(required);
    }
    item.setTitle(title)
        .setHelpText(Description);
}


// test

function testReplaceConfVars() {
    Logger.log(replaceConfVars("test [SECTION] [SECTION] [nothing] BLBx A"));
}