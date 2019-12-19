//Static variables --------------------------------------------------------------------------------------------------------------------------------------------------------//
var CONFIRM_MAIL = "confirm sent";
var REGISTER_MAIL = "register sent";
var EXTRA_MAIL = "extra sent";

//sends confirmation email to the participant in row --------------------------------------------------------------------------------------------------------------------------------------------------------//
function sendconfirmationEmail(row) {
  if (!options["AUTO_CONF_MAIL"]) {
    return;
  }
  var scriptRange = row.getCell(1, indexOfScript);
  var emailSent = scriptRange.getValue();
  var email = getByNameRow("Email", row);
  if (emailSent != CONFIRM_MAIL && email != "" && getByNameRow("Paid", row) == "yes") {  // Prevents sending duplicates
    var subject = "Confirmation " + options["EVENT_TITLE"];
    //if the user has a whole confirm draft
    if (sendGmailTemplate(email, subject, row, {}, options["CONF_MAIL_NAME"])) {
      scriptRange.setValue(CONFIRM_MAIL);
    }
  }
}
//sends a registration email to the participant in row --------------------------------------------------------------------------------------------------------------------------------------------------------//
function sendRegisterEmail(row) {
  var script_registration_mail_name = options["REG_MAIL_NAME"];
  if (!options["AUTO_REG_MAIL"] || script_registration_mail_name == "") {
    return;
  }
  //var scriptRange = registerSheet.getRange(row + 1, indexOfScript, 1, 1);
  var scriptRange = row.getCell(1, indexOfScript);
  var email = getByNameRow("Email", row);
  var emailSent = scriptRange.getValue();     // column where we can check if the user already got an email
  if (emailSent != REGISTER_MAIL && emailSent != CONFIRM_MAIL && email != "") {  // Prevents sending duplicates
    var subject = "Registration " + options["EVENT_TITLE"];
    //if the user has a whole confirm draft
    if (options["REG_MAIL_PAYMENTTYPE"]) {
      var paymentmethod = getByNameRow("Payment method", row);
      script_registration_mail_name = script_registration_mail_name + "_" + paymentmethod;
    }
    console.info('Sending reg mail %s running as', email, getUserEmail()); //DEBUG
    if (sendGmailTemplate(email, subject, row, {}, script_registration_mail_name)) {
      scriptRange.setValue(REGISTER_MAIL);
      SpreadsheetApp.flush();
    } else {
      Logger.log("error while sending mail");
    }
    // Make sure the cell is updated right away in case the script is interrupted
  }
}

function getUserEmail() {
  var userEmail = PropertiesService.getUserProperties().getProperty("userEmail");
  if(!userEmail) {
    var protection = SpreadsheetApp.getActive().getRange("A1").protect();
    // tric: the owner and user can not be removed
    protection.removeEditors(protection.getEditors());
    var editors = protection.getEditors();
    if(editors.length === 2) {
      var owner = SpreadsheetApp.getActive().getOwner();
      editors.splice(editors.indexOf(owner),1); // remove owner, take the user
    }
    userEmail = editors[0];
    protection.remove();
    // saving for better performance next run
    // PropertiesService.getUserProperties().setProperty("userEmail",userEmail);
  }
  return userEmail;
}


//sends an email to all participants when clicking extra email button --------------------------------------------------------------------------------------------------------------------------------------------------------//
function sendExtraEmails() {
  var script_extra_mail_name = options["EXTRA_MAIL_NAME"];
  if (script_extra_mail_name == "") {
    showAlert("no extra mail name", "specify extra mail name in options sheet");
    return;
  }
  
  for (var i = 2; i <= registerSheet.getLastRow(); ++i) {
    var row = registerSheet.getRange(i, 1, registerSheet.getLastRow(), registerHeaders.length); // let it read more 
    
    var emailSent = row.getCell(1, indexOfScript);
    var email = row.getCell(1, getColumnId("Email")).getValue();
    if (emailSent.getValue() != EXTRA_MAIL && email != "") {  // Prevents sending duplicates
      var subject = options["EVENT_TITLE"];
      
      if (options["EXTRA_PAID"]) {
        if (row.getCell(1, indexOfPaid).getValue() == "yes") {
          if (sendGmailTemplate(email, subject, row, {}, options["EXTRA_MAIL_NAME"])) {
            emailSent.setValue(EXTRA_MAIL);
          }
        }
      } else {
        if (sendGmailTemplate(email, subject, row, {}, options["EXTRA_MAIL_NAME"])) {
          emailSent.setValue(EXTRA_MAIL);
        }
      }
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}

/**
* Insert the given email body text into an email template, and send
* it to the indicated recipient. The template is a draft message with
* the subject "TEMPLATE"; if the template message is not found, an
* exception will be thrown. The template must contain text indicating
* where email content should be placed: {BODY}.
*
* @param {String} recipient  Email address to send message to.
* @param {String} subject    Subject line for email.
* @param {String} body       Email content, may be plain text or HTML.
* @param {Object} options    (optional) Options as supported by GmailApp.
*
* @returns        GmailApp   the Gmail service, useful for chaining
*/
function sendGmailTemplate(recipient, subject, row, mail_options, draftsubject) {
  mail_options = mail_options || {};  // default is no options
  
  if (!existsDraft(draftsubject)) throw new Error("TEMPLATE not found in drafts folder: " + draftsubject);
  var template = readDraft(draftsubject)
  
  // Generate htmlBody from template, with provided text body
  var imgUpdates = updateInlineImages(template);
  var message = imgUpdates.templateBody;
  message = replaceTerms(message, row);
  mail_options.htmlBody = message;
  mail_options.attachments = imgUpdates.attachments;
  mail_options.inlineImages = imgUpdates.inlineImages;
  var body = message;
  return GmailApp.sendEmail(recipient, subject, body, mail_options);
}

//fixes any attachments or pictures in email template --------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
* @param   {GmailMessage} template  Message to use as template
* @returns {Object}                 An object containing the updated
*                                   templateBody, attachments and inlineImages.
*/
function updateInlineImages(template) {
  //////////////////////////////////////////////////////////////////////////////
  // Get inline images and make sure they stay as inline images
  //////////////////////////////////////////////////////////////////////////////
  var templateBody = template.getBody();
  var rawContent = template.getRawContent();
  var attachments = template.getAttachments();
  var regMessageId = new RegExp(template.getId(), "g");
  if (templateBody.match(regMessageId) != null) {
    var inlineImages = {};
    var nbrOfImg = templateBody.match(regMessageId).length;
    var imgVars = templateBody.match(/<img[^>]+>/g);
    var imgToReplace = [];
    if (imgVars != null) {
      for (var i = 0; i < imgVars.length; i++) {
        if (imgVars[i].search(regMessageId) != -1) {
          var id = imgVars[i].match(/realattid=([^&]+)&/);
          if (id != null) {
            var temp = rawContent.split(id[1])[1];
            temp = temp.substr(temp.lastIndexOf('Content-Type'));
            var imgTitle = temp.match(/name="([^"]+)"/);
            if (imgTitle != null) imgToReplace.push([imgTitle[1], imgVars[i], id[1]]);
          }
        }
      }
    }
    for (var i = 0; i < imgToReplace.length; i++) {
      for (var j = 0; j < attachments.length; j++) {
        if (attachments[j].getName() == imgToReplace[i][0]) {
          inlineImages[imgToReplace[i][2]] = attachments[j].copyBlob();
          attachments.splice(j, 1);
          var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
          templateBody = templateBody.replace(imgToReplace[i][1], newImg);
        }
      }
    }
  }
  var updatedTemplate = {
    templateBody: templateBody,
    attachments: attachments,
    inlineImages: inlineImages
  }
  return updatedTemplate;
}
//Replace Terms in Email --------------------------------------------------------------------------------------------------------------------------------------------------------//
function replaceTerms(message, row) {
  var columnrange = registerSheet.getRange(1, 1, 1, registerSheet.getLastColumn());
  var columnvalues = columnrange.getValues();
  
  message = replaceTemplateKey(message, "PAYPAL",calculatePayPal(row));
  message = replaceConfVars(message);
  message = replaceTemplateKey(message, "PRICE", calculatePrice(row));
  
  
  for (var j = 0; j < columnvalues[0].length; j++) {
    var key = columnvalues[0][j];
    if (key != null) {
      message = replaceTemplateKey(message, key, getByNameRow(columnvalues[0][j], row));
    }
  }
  return message;
}
/**
* creates the draft mails according to the names in the fields 
*/
function createDraftMails() {
  var script_registration_mail_name = options["REG_MAIL_NAME"];
  var script_confirm_mail_name = options["CONF_MAIL_NAME"];
  var script_extra_mail_name = options["EXTRA_MAIL_NAME"];
  
  var payment_methods = [];
  var questions = getAllQuestions();
  for each(var question in questions) {
    if (question.title == "Payment method") {
      payment_methods = question.options;
    }
  }
  
  if (script_confirm_mail_name != "" && !existsDraft(script_confirm_mail_name)) {
    GmailApp.createDraft("", script_confirm_mail_name, "", { htmlBody: "Confirm email" });
  }
  if (script_extra_mail_name != "" && !existsDraft(script_extra_mail_name)) {
    GmailApp.createDraft("", script_extra_mail_name, "", { htmlBody: "Extra email" });
  }
  if (script_registration_mail_name != "") {
    if (options["REG_MAIL_PAYMENTTYPE"]) {
      // payment specific email
      for each(var payment in payment_methods) {
        var subject = script_registration_mail_name + "_" + payment;
        if (!existsDraft(subject)) {
          GmailApp.createDraft("", subject, "", { htmlBody: "Fill Payment info " + payment });
        }
      }
    } else {
      if (!existsDraft(script_registration_mail_name)) {
        GmailApp.createDraft("", script_registration_mail_name, "", { htmlBody: "Registration email " });
      }
    }
  }
}
/**
* check if hte given subject exists in the list of draft mails
* drafts = GmailApp.getDraftMessages();
* */
function existsDraft(subject) {
  drafts = getDraftMessages();
  for each(var draft in drafts) {
    if (draft.getSubject() == subject) {
      return true;
    }
  }
  return false;
}



function readDraft(draft_name) {
  drafts = getDraftMessages();
  for (var i = 0; i < drafts.length; ++i) {
    if (drafts[i].getSubject() == draft_name) {
      return drafts[i];
    }
  }
}

function replaceTemplateVars(html_body) {
  var replaced = "";
  return replaced;
}
function sendMail(email_to, subject, html_body) {
  MailApp.sendEmail({
    to: email_to,
    subject: subject,
    htmlBody: html_body
  });
}
