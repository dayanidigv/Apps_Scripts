function doPost(e) {
  try {
    
    var name = e.parameter.name;
    var email = e.parameter.email;
    var subject = e.parameter.subject;
    var message = e.parameter.message;
    var phoneno = e.parameter.phoneno;
 
    // Set up the email details
    var recipient = "<gmail>";
    var emailSubject = "Enquire From <Website-name> Website: " + subject;
    var emailMessage = `
    Name: ${name}
    Email: ${email}
    Subject: ${subject}
    Message: ${message}
    Phone Number: ${phoneno}

    Click here to reply to this message 👇:
    https://mail.google.com/mail/u/0/?tf=cm&fs=1&to=${encodeURIComponent(email)}&hl=en&su=${encodeURIComponent("Reply to " + name + "'s Enquiry")}&body=${encodeURIComponent(`Hi ${name},\n\nThank you for your message!\n\nHere's our response:\n\n`)}
`;

    // Send the email using MailApp
    MailApp.sendEmail({
      to: recipient,
      subject: emailSubject,
      body: emailMessage
    });

    // Return a success response
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "Email sent successfully"})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Log the error and return a failure response
    Logger.log("Error: " + error.message);
    return ContentService.createTextOutput(JSON.stringify({"status": "failure", "message": "Failed to send email", "error": error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
