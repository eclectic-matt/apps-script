/* Checks quotas with try-catch and sends custom email */
function sendEmailCheckQuotas(){

  /* EDIT THESE */
  var emailToSendTo = "example@foo.bar";
  var subject = "Email Subject";
  var message = "<html><body>Put your message body in here!</body></html>";
  var fromName = "From this Place!";

  /* TRY-CATCH FOR EMAIL QUOTA LIMITS */
  try {
    MailApp.sendEmail(emailToSendTo, subject, message, { name: fromName, htmlBody: message });
  }catch(e){
      Browser.msgBox("Sorry, an error occurred sending the update email: \n\r" + e + "\n\rThe email text will be shown in a browser pop-up (please copy and paste)");
      Browser.msgBox(message);
  }

}
