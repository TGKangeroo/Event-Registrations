

// checks and logs issues with event fields that could cause a crash of the event scripts -------------------------------------------------------------------------------------------------------------
function checkEventInformation() {
  var check=true; 
  if(getFieldValue('event_title') ==""){check=false;Logger.log('Fill in the event title');}
  if(getFieldValue('event_start_date')==''){check=false;Logger.log('Fill in the event start date');}
  if(getFieldValue('event_start_time')==''){check=false;Logger.log('Fill in the event start time');}
  if(getFieldValue('event_end_date')==''){check=false;Logger.log('Fill in the event end date');}
  if(getFieldValue('event_end_time')==''){check=false;Logger.log('Fill in the event end time');}
  if(getFieldValue('event_meetingpoint')==''){check=false;Logger.log('Fill in the event meetingpoint');}
  if( getFieldValue('event_max_participants')=='' && getFieldValue('event_max_participants')!=0 ){check=false;Logger.log('Fill in the event maximum amount of participants');}
  if(getFieldValue('event_isPaid')==''){check=false;Logger.log('Fill in if the event is a paid event or not');}
  if(getFieldValue('event_section')==''){check=false;Logger.log('Fill in the event section');}
  if(getFieldValue('cancellation_policy')==''){check=false;Logger.log('Fill in the cancellation policy');}
  if(getFieldValue('dsgvo_policy')==''){check=false;Logger.log('Fill in the DSGVO Policy');}
  if(getFieldValue('script_registration_close_date')==''){check=false;Logger.log('Fill in the registration closing date');}
 
  return check;
}
// checks and logs issues with question fields that could cause a crash of the event scripts -------------------------------------------------------------------------------------------------------------
function checkQuestions(){
var check=true;
var checkEmail=false;
var checkSurName=false;
var checkFirstName=false;
var checkPaymentMethod=false;
  
  if(getFieldValue('script_register_on_pay')!="yes"){checkPaymentMethod=true;}  
  
  
var questions = getAllQuestions();
  
  // array with all current question types
var questiontypes=['text','email','radiobutton','date','checkbox','dropdown','time','datetime','duration'];
  
  // array with all question types that need the option field
var questionOptionTypes=  ['radiobutton','checkbox','dropdown'];
  
  for(var i=0;i<questions.length;i++){  
  if( questions[i][0] !="" ){
  if( questions[i][1] ==""){check=false;Logger.log('question ' + i + ' is missing a type');}
  if( questions[i][3] !=false && questions[i][3] !=true){check=false;Logger.log('question ' + i + ' is missing required');}  
  if( questionOptionTypes.indexOf(questions[i][1]) != -1 && questions[i][2] =="") {check=false;Logger.log('question ' + i + ' is missing options');}  
    
    if(questions[i][0] =="First name"){checkFirstName=true;}
    if(questions[i][0] =="Surname"){checkSurName=true;}
    if(questions[i][0] =="Email"){checkEmail=true;}
    if(getFieldValue('script_register_on_pay') == "yes" && questions[i][0] == "Payment method"){checkPaymentMethod=true;}
  }
  }
  
  if(checkFirstName != true || checkSurName!=true || checkEmail !=true || checkPaymentMethod !=true){check=false;}
 
  if(checkFirstName !=true){Logger.log('You do not have a question called First name');}
  if(checkSurName !=true){Logger.log('You do not have a question called Surname');}
  if(checkEmail !=true){Logger.log('You do not have a question called Email');}
  if(checkPaymentMethod !=true){Logger.log('You do not have a question called Payment Method, but you devide emails based on payment');}
  
  return check;
}


// checks and logs issues with price fields that could cause a crash of the event scripts -------------------------------------------------------------------------------------------------------------
function checkPrices(){
var check=true;
var questions = getAllQuestionNames();
var prices = getAllPrices();
  
  for (var i=0;i<prices.length;i++){
    if( prices[i][0] != ""){
    
      if(getFieldValue('event_isPaid')=="no" && prices[i][0] !=""){check=false;Logger.log('Choose if there is payment needed for the event');}
      if(prices[i][1] == ""){check=false;Logger.log('price ' + i + ' is missing a value');}
      if(prices[i][3] ==""){check=false;Logger.log('price ' + i + ' is missing a corresponding field');}
      if(prices[i][3] != "Base Price" && prices[i][4]==""){check=false;Logger.log('price ' + i + ' is missing an answer');}
      if(prices[i][3] !="Base Price" && questions.indexOf(prices[i][3]) < 0){check=false;Logger.log('price ' + i + ' is missing some information');}
    
    }
  
  }
 
  
return check;
}

// checks and logs issues with script fields that could cause a crash of the event scripts -------------------------------------------------------------------------------------------------------------
function checkScriptingOptions(){
var check=true;
  
 
if(getFieldValue('script_sticky_names') ==""){check=false;Logger.log('Choose if you want sticky names or not');}
  if(getFieldValue('script_close_form_max_part') ==""){check=false;Logger.log('Choose if the registrations should close when max participants has been reached');}
  if(getFieldValue('script_color_on_paid') ==""){check=false;Logger.log('Choose if you want to change the color of paid participants');}
  if(getFieldValue('script_extra_mail_on_pay') ==""){check=false;Logger.log('Choose if the extra email should only be sent to paying participants');}
  if(getFieldValue('script_auto_confirm_mails') ==""){check=false;Logger.log('Choose if confirmation emails should be sent automatically');}
  if(getFieldValue('script_auto_confirm_mails')== "yes" && getFieldValue('script_confirm_mail_name') ==""){check=false;Logger.log('Fill in the confirmation email draft subject');}
  if(getFieldValue('script_auto_registration_mails') ==""){check=false;Logger.log('Choose if registration emails should be sent automatically');}
  if( getFieldValue('script_auto_registration_mails')== "yes" && getFieldValue('script_registration_mail_name') ==""){check=false;Logger.log('Fill in the registration email draft subject');}
  if(getFieldValue('script_register_on_pay') ==""){check=false;Logger.log('Choose if different emails should be sent based on payment type');}

return check;
}

// checks and logs issues with the finance fields that could cause a crash of the event scripts -------------------------------------------------------------------------------------------------------------
function checkFinances(){
  var check=true;
  if (getFinanceValue('paypal_allowed') == "yes"){
  if(getFinanceValue('paypal_email_address')==''){check=false;Logger.log('The PayPal email address is empty');}
  if(getFinanceValue('paypal_business_percentage')==''){check=false;Logger.log('The Business percentage is empty');}
    if(getFinanceValue('paypal_description')==''){check=false;Logger.log('The PayPal description is empty');}
  
  }
  if (getFinanceValue('banktransfer_allowed') == "yes"){
   if(getFinanceValue('banktransfer_account_owner')==''){check=false;Logger.log('The bank account owner is empty');}
     if(getFinanceValue('banktransfer_bank_name')==''){check=false;Logger.log('The bank name is empty');}
     if(getFinanceValue('banktransfer_iban')==''){check=false;Logger.log('The Iban is empty');}
     if(getFinanceValue('banktransfer_bic')==''){check=false;Logger.log('The BIC is empty');}
     if(getFinanceValue('banktransfer_description')==''){check=false;Logger.log('The bank transfer description is empty');}
  
  }
  if (getFinanceValue('cash_allowed') == "yes"){
   if(getFinanceValue('cash_office_address')==''){check=false;Logger.log('The office address is empty');}
    if(getFinanceValue('cash_office_days')==''){check=false;Logger.log('The office days are empty');}
    if(getFinanceValue('cash_office_hours')==''){check=false;Logger.log('The office hours are empty');}
  
  }
 
return check;
}

// checks and logs issues with the event script already generating parts of its output before -------------------------------------------------------------------------------------------------------------
function checkPrerequisits(){
var check= true;
  
  if(getFieldValue('printlist_added')!=""){check=false;Logger.log('Print list has already been added');}
  if(getFieldValue('rows_added')!=""){check=false;Logger.log('Registrations rows have already been added');}
  if(getFieldValue('form_made')!=""){check=false;Logger.log('There is already a form made');}
  
 
return check;
}

// Executes all the previous checks -------------------------------------------------------------------------------------------------------------
function checkEverything(){
var check = true;
  
  if(checkFinances()!=true){check=false;}
  if(checkScriptingOptions()!=true){check=false;}
  if(checkPrices()!=true){check=false;}
  if(checkQuestions()!=true){check=false;}
  if(checkEventInformation()!=true){check=false;}
  if(checkPrerequisits()!=true){check=false;}
  
  Logger.clear();
return check;
}


function readyForm(){
  if(checkEverything()){
  makeForm();
  }else{
  
  showAlert('You did something wrong','Please make sure all boxes appear green before making a form!');
  }

}

