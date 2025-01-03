function doPost(e) {
  const data = e.parameter;
  GmailApp.sendEmail(
    data.email, 
    "Form Submission Received", 
    `Hi ${data.name}, we got your submission!`
  );
  return ContentService.createTextOutput(
    'Form submitted & email sent!'
  );
}
