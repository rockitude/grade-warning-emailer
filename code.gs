//Load the menu
function onOpen(){
  loadMenu(); 
}

function loadMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [];
  menuItems.push({name: "Send Emails", functionName: "sendEmails"});
  ss.addMenu("Grade Warning Emailer", menuItems);
}

//Send the Emails
function sendEmails() {
  
  //initialize and populate variables
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange("B2:E30");
  var data = range.getValues();
  var studentsEmailed = [];
  var parentsEmailed = [];
  
  
  //set email subject line to value in spreadsheet. consider adding a warning later in case they don't fill the cell in.
  var subSheet = ss.getSheetByName("Email Text");
  var subject = subSheet.getRange("A7").getValue(); //"Weekly F-Watch Email for Mr. Lipson's Class";
  var teacherName = subSheet.getRange("A5").getValue(); //Teacher's Name
  
  for (var i = 0; i < 30; i++){
    //set variables
    var row = data[i];
    var emailAddress = row[1];
    var parentEmailAddress = row[2];
    var grade = row[3];
    
    //if Name, Student Email Address and Grade aren't empty...
    if(row[0] !== "" && row[1] !== "" && row[3] !== ""){
    
      //email student if necessary, return true or false
      if(emailAddress !== ""){
        var emailSent = sendStudentEmail(emailAddress, grade, subject, generateBody(row[0], row[3], teacherName));
        if(emailSent) {studentsEmailed.push(emailAddress);}
      }
    
      //email parent if necessary, return true or false
      if(parentEmailAddress !== ""){
        var parentEmailSent = sendParentEmail(emailAddress, grade, subject, generateBody(row[0], row[3], teacherName));
        if(parentEmailSent){parentsEmailed.push(parentEmailAddress);}
      }
  }
  else{
    break;
  }
}
  
  //report how many students and parents were mailed in a popup
  var status_html = studentsEmailed.length + " student(s) emailed.\n";
  status_html += parentsEmailed.length + " parent(s) emailed.";
  Browser.msgBox(status_html);

}

function generateBody(studentName, grade, teacherName){
  var body = "<p><b>Attention " + studentName + " and Parents:</b></p>";
  body += "<p>Your grade in ";
  body += teacherName;
  body += "\'s class is currently: " + grade + "</p>";
  
  body += SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Text").getRange("A10").getValue();
  
  return body;
}

function sendStudentEmail(emailAddress, grade, subject, htmlBody){
  if(grade < 60){
    MailApp.sendEmail(emailAddress, subject, "", {htmlBody: htmlBody});
    return true;
  }
  else{
    return false;
  }
}

function sendParentEmail(parentEmailAddress, grade, subject, htmlBody){
  if(grade < 60){
    MailApp.sendEmail(parentEmailAddress, subject, "", {htmlBody: htmlBody});
    return true;
  }
  else{
    return false;
  }
}
