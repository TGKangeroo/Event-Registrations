var questions = [];
function getAllQuestions() {
    var values = ss.getRange("Questions!A3:F50").getValues();
    
    for (var row in values) {

        var enabled = values[row][0];
        if(enabled == true) {
            var form_type = values[row][3];
            var form_options = [];
            if(form_type == "radiobutton" || form_type == "dropdown"|| form_type == "checkbox" ) {
                form_options = values[row][4].split(",");
            }
            questions.push( { title: values[row][1], required: values[row][2], type: form_type, options: form_options, desc: values[row][5] } );
        }
    }
    return questions;
}

