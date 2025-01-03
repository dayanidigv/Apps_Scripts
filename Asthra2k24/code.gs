function reloadAllData() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var sheetNames = getAllSheetNames();

  sheetNames.forEach(function (name) {
    if (name != "Sheet1" && name != "Dashboard") {
      sheets.deleteSheet(sheets.getSheetByName(name));
    }
  });

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") {
      return;
    }

    if (firstEvent) {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent) { // Check if second event is different from first
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reloadDashboardData() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();
  var isSheet = sheets.getSheetByName("Dashboard");

  var dashboardCounts = {
    "Total No of Registrations": 0,
    "Total Registrations Amount": 0,
    "Total No of Participants": 0,
    "Total Non-veg": 0,
    "Total Veg": 0,
    "AR VR Workshop": 0,
    "Just a Minute (JAM)": 0,
    "Paper Presentation": 0,
    "No of Paper Presentation Participants": 0,
    "Mobile Photography": 0,
    "Poster Design": 0,
    "Web Design": 0,
    "Project Expo": 0,
    "No of Paper Project Expo": 0,
  };

  if (isSheet) {
    sheets.deleteSheet(isSheet);
  }

  data.forEach(function (entry) {
    var foodType = entry[7];
    var firstEvent = entry[10];
    var secondEvent = entry[11];
    var amount = entry[12];

    if (firstEvent === "1st Event") {
      return;
    } else {
      dashboardCounts["Total No of Registrations"] += 1;
    }

    if (amount) {
      dashboardCounts["Total Registrations Amount"] += parseInt(amount.replace("₹ ", ""), 10);
    }

    if (firstEvent) {
      dashboardCounts[firstEvent] += 1;
    }
    if (secondEvent && secondEvent !== firstEvent) {
      dashboardCounts[secondEvent] += 1;
    }

    if (firstEvent === "Paper Presentation" || secondEvent === "Paper Presentation") {
      var ppTeamSize = entry[15];

      if (ppTeamSize !== "") {
        dashboardCounts["Total No of Participants"] += parseInt(ppTeamSize, 10);
        dashboardCounts["No of Paper Presentation Participants"] += parseInt(ppTeamSize, 10);

        if (foodType) {
          if (foodType === "Non-Veg") {
            dashboardCounts["Total Non-veg"] += parseInt(ppTeamSize, 10);
          } else {
            dashboardCounts["Total Veg"] += parseInt(ppTeamSize, 10);
          }
        }
      }
    } else if (firstEvent === "Project Expo" || secondEvent === "Project Expo") {
      var peTeamSize = entry[30];
      if (peTeamSize !== "") {
        dashboardCounts["Total No of Participants"] += parseInt(peTeamSize, 10);
        dashboardCounts["No of Paper Project Expo"] += parseInt(peTeamSize, 10);
        if (foodType) {
          if (foodType === "Non-Veg") {
            dashboardCounts["Total Non-veg"] += parseInt(peTeamSize, 10); // Add to Total Non-veg if foodType is Non-Veg
          } else {
            dashboardCounts["Total Veg"] += parseInt(peTeamSize, 10); // Add to Total Veg if foodType is not Non-Veg
          }
        }
      }
    } else {
      dashboardCounts["Total No of Participants"] += 1;
      if (foodType) {
        if (foodType === "Non-Veg") {
          dashboardCounts["Total Non-veg"] += 1;
        } else {
          dashboardCounts["Total Veg"] += 1;
        }
      }
    }

  });


  var dashboardSheet = getOrCreateSheetDashboard(sheets, "Dashboard");
  appendDashboardDataToSheet(dashboardSheet, dashboardCounts);
}

function reload_all_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Final - list");

  if (isSheet) {
    sheets.deleteSheet(isSheet);
  }

  let sheet = getOrCreateSheet(sheets, "Final - list");

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    appendDataToSheet(sheet, entry);

    if (firstEvent === "Paper Presentation" || secondEvent === "Paper Presentation") {
      var ppTeamSize = entry[15];
      if (ppTeamSize == 2) {
        appendDataToSheet(sheet, entry.slice(15, 22));
      } else if (ppTeamSize == 3) {
        appendDataToSheet(sheet, entry.slice(15, 22));
        appendDataToSheet(sheet, entry.slice(22, 29));
      }
    }

    

    if (firstEvent === "Project Expo" || secondEvent === "Project Expo") {
      var peTeamSize = entry[20];
      if (peTeamSize == 2) {
        appendDataToSheet(sheet, entry.slice(30, 37));
      } 
    }

  });
}

function reload_ARVR_final_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("AR VR Workshop - final");

  if (isSheet) {
    sheets.deleteSheet(isSheet);
  }

  let sheet = getOrCreateSheet(sheets, "AR VR Workshop - final");

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "AR VR Workshop" || secondEvent === "AR VR Workshop") {
      appendDataToSheet(sheet, entry);
    }

    if (firstEvent === "Paper Presentation" || secondEvent === "Paper Presentation") {
      if (firstEvent === "AR VR Workshop" || secondEvent === "AR VR Workshop") {
        var ppTeamSize = entry[15];
        if (ppTeamSize == 2) {
          appendDataToSheet(sheet, entry.slice(15, 22));
        } else if (ppTeamSize == 3) {
          appendDataToSheet(sheet, entry.slice(15, 22));
          appendDataToSheet(sheet, entry.slice(22, 29));
        }
      }
    }

  });
}

function reload_ARVR_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("AR VR Workshop");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("AR VR Workshop"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "AR VR Workshop") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "AR VR Workshop") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_JAM_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Just a Minute (JAM)");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Just a Minute (JAM)"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Just a Minute (JAM)") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Just a Minute (JAM)") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_Paper_Presentation_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Paper Presentation");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Paper Presentation"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Paper Presentation") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Paper Presentation") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_Mobile_Photography_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Mobile Photography");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Mobile Photography"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Mobile Photography") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Mobile Photography") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_Poster_Design_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Poster Design");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Poster Design"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Poster Design") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Poster Design") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_Web_Design_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Web Design");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Web Design"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Web Design") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Web Design") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function reload_Project_Expo_Data() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheets.getDataRange().getValues();

  var isSheet = sheets.getSheetByName("Project Expo");

  if (isSheet) {
    sheets.deleteSheet(sheets.getSheetByName("Project Expo"));
  }

  data.forEach(function (entry) {
    var firstEvent = entry[10];
    var secondEvent = entry[11];

    if (firstEvent === "1st Event") { return; }

    if (firstEvent === "Project Expo") {
      var sheet = getOrCreateSheet(sheets, firstEvent);
      appendDataToSheet(sheet, entry);
    }

    if (secondEvent && secondEvent !== firstEvent && secondEvent === "Project Expo") {
      var sheet = getOrCreateSheet(sheets, secondEvent);
      appendDataToSheet(sheet, entry);
    }

  });
}

function getOrCreateSheet(sheets, name) {
  var sheet = sheets.getSheetByName(name);

  if (!sheet) {
    sheet = sheets.insertSheet(name);
    sheet.appendRow(["Name", "Email", "Phone", "College", "Department", "Year"]);
  }
  return sheet;
}

function getAllSheetNames() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetNames = [];

  for (var i = 1; i < sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }

  return sheetNames;
}

function appendDataToSheet(sheet, entry) {
  var rowData = [
    entry[1], // Name
    entry[2], // Email
    entry[3], // Phone
    entry[4], // College
    entry[5], // Department
    entry[6], // Year
  ];
  sheet.appendRow(rowData);
}

function getOrCreateSheetDashboard(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function appendDashboardDataToSheet(sheet, data) {
  var headers = Object.keys(data);
  headers.forEach(function (header) {
    sheet.appendRow([header, data[header]]);
  });
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var jsonData = convertDataToJson(data);
  return ContentService.createTextOutput(jsonData).setMimeType(ContentService.MimeType.JSON);
}

function convertDataToJson(data) {
  var headers = data[0];
  var jsonData = [];

  // Loop through rows starting from the second row (index 1)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entry = {};

    // Loop through columns
    for (var j = 0; j < headers.length; j++) {
      entry[headers[j]] = row[j];
    }

    jsonData.push(entry);
  }

  return JSON.stringify(jsonData);
}

function doPost(e) {
  try {
    var formData = JSON.parse(e.postData.contents);
    handleFormSubmission(formData);
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'message': "Registration details received successfully." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error processing form submission:", error);
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'message': "Error processing form submission." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

async function handleFormSubmission(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = [
      new Date(),
      formData.basicInfo.name,
      formData.basicInfo.email,
      formData.basicInfo.phone,
      formData.basicInfo.college,
      formData.basicInfo.department,
      formData.basicInfo.year,
      formData.basicInfo.foodType,
    ];

    var idCardFolder = createFolderIfNotExists('IdCard');
    var idCardFileUrl = createFileFromBase64(formData.basicInfo.idCard, idCardFolder);

    row.push(idCardFileUrl ? createImageFormula(idCardFileUrl) : "");

    row.push(formData.eventInfo.noOfEvents);
    if (formData.eventInfo.noOfEvents == "1") {
      row.push(formData.eventInfo.events[0]);
      row.push("");
    } else {
      row.push(formData.eventInfo.events[0]);
      row.push(formData.eventInfo.events[1]);
    }

    row.push(formData.payment.amount);

    var transactionScreenshotFolder = createFolderIfNotExists('TransactionScreenshot');
    var transactionScreenshotFileUrl = createFileFromBase64(formData.payment.transactionScreenshot, transactionScreenshotFolder);

    row.push(transactionScreenshotFileUrl ? createImageFormula(transactionScreenshotFileUrl) : "");
    row.push(formData.payment.transactionID);

    row = handleTeamSection(row, formData.paperPresentation, idCardFolder);
    row = handleTeamSection(row, formData.projectExpo, idCardFolder);
    row = await sendRegistrationSuccessEmail(row, formData.basicInfo.email, formData.basicInfo.name);
    sheet.appendRow(row);
  } catch (error) {
    console.error("Error handling form submission:", error);
    throw new Error("Error handling form submission.");
  }
}

function handleTeamSection(row, teamSection, idCardFolder) {
  try {
    var teamSize = teamSection.teamSize;
    row.push(teamSize || "");
    teamSection.teamMembers.forEach((member, index) => {
      var idCardFileUrl = member.idCard != null ? createFileFromBase64(member.idCard, idCardFolder) : null;
      var memberRow = [
        member.name,
        member.email,
        member.phone,
        member.college,
        member.department,
        member.year,
        idCardFileUrl ? createImageFormula(idCardFileUrl) : "",
      ];
      row = row.concat(memberRow);
    });
    return row;
  } catch (error) {
    console.error(`Error handling ${teamSectionName} team section:`, error);
    throw new Error(`Error handling ${teamSectionName} team section.`);
  }
}

function createImageFormula(url) {
  return `=IMAGE("${url}", 1)`;
}

function createFolderIfNotExists(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function createFileFromBase64(obj, folderName) {
  try {
    var decoded = Utilities.base64Decode(obj.base64);
    var folder = createFolderIfNotExists(folderName);
    var blob = Utilities.newBlob(decoded, obj.type, obj.fileName);
    var newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var link = newFile.getDownloadUrl();
    return link;
  } catch (error) {
    throw new Error("Error creating or sharing the file: " + error.message);
  }
}

async function sendRegistrationSuccessEmail(row, recipient, participantName) {
  var subject = "Welcome to ASTHRA 2K24 - Registration Confirmation";
  var greeting = "Dear Participant,";
  var intro = "We are thrilled to welcome you to ASTHRA 2K24!";
  var message = "Your registration has been successfully recorded, and our team is currently processing your verification. This may take a little time, but rest assured, we'll notify you as soon as it's completed.";
  var closing = "Thank you for choosing ASTHRA 2K24. Should you have any questions or need assistance, please feel free to reach out to our support team. Get ready for an exciting symposium experience!";

  var formattedBody = `<!DOCTYPE html
    PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en"
    style="font-family:arial, 'helvetica neue', helvetica, sans-serif">

    <head>
        <meta charset="UTF-8">
        <meta content="width=device-width, initial-scale=1" name="viewport">
        <meta name="x-apple-disable-message-reformatting">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta content="telephone=no" name="format-detection">
        <title>Registration Confirmation</title>
        <link href="https://fonts.googleapis.com/css2?family=Imprima&display=swap" rel="stylesheet"> <!--<![endif]-->
        <style type="text/css">
            #outlook a {
                padding: 0;
            }

            .es-button {
                mso-style-priority: 100 !important;
                text-decoration: none !important;
            }

            a[x-apple-data-detectors] {
                color: inherit !important;
                text-decoration: none !important;
                font-size: inherit !important;
                font-family: inherit !important;
                font-weight: inherit !important;
                line-height: inherit !important;
            }

            .es-desk-hidden {
                display: none;
                float: left;
                overflow: hidden;
                width: 0;
                max-height: 0;
                line-height: 0;
                mso-hide: all;
            }

            @media only screen and (max-width:600px) {

                p,
                ul li,
                ol li,
                a {
                    line-height: 150% !important
                }

                h1,
                h2,
                h3,
                h1 a,
                h2 a,
                h3 a {
                    line-height: 120%
                }

                h1 {
                    font-size: 30px !important;
                    text-align: left
                }

                h2 {
                    font-size: 24px !important;
                    text-align: left
                }

                h3 {
                    font-size: 20px !important;
                    text-align: left
                }

                .es-header-body h1 a,
                .es-content-body h1 a,
                .es-footer-body h1 a {
                    font-size: 30px !important;
                    text-align: left
                }

                .es-header-body h2 a,
                .es-content-body h2 a,
                .es-footer-body h2 a {
                    font-size: 24px !important;
                    text-align: left
                }

                .es-header-body h3 a,
                .es-content-body h3 a,
                .es-footer-body h3 a {
                    font-size: 20px !important;
                    text-align: left
                }

                .es-menu td a {
                    font-size: 14px !important
                }

                .es-header-body p,
                .es-header-body ul li,
                .es-header-body ol li,
                .es-header-body a {
                    font-size: 14px !important
                }

                .es-content-body p,
                .es-content-body ul li,
                .es-content-body ol li,
                .es-content-body a {
                    font-size: 14px !important
                }

                .es-footer-body p,
                .es-footer-body ul li,
                .es-footer-body ol li,
                .es-footer-body a {
                    font-size: 14px !important
                }

                .es-infoblock p,
                .es-infoblock ul li,
                .es-infoblock ol li,
                .es-infoblock a {
                    font-size: 12px !important
                }

                *[class="gmail-fix"] {
                    display: none !important
                }

                .es-m-txt-c,
                .es-m-txt-c h1,
                .es-m-txt-c h2,
                .es-m-txt-c h3 {
                    text-align: center !important
                }

                .es-m-txt-r,
                .es-m-txt-r h1,
                .es-m-txt-r h2,
                .es-m-txt-r h3 {
                    text-align: right !important
                }

                .es-m-txt-l,
                .es-m-txt-l h1,
                .es-m-txt-l h2,
                .es-m-txt-l h3 {
                    text-align: left !important
                }

                .es-m-txt-r img,
                .es-m-txt-c img,
                .es-m-txt-l img {
                    display: inline !important
                }

                .es-button-border {
                    display: block !important
                }

                a.es-button,
                button.es-button {
                    font-size: 18px !important;
                    display: block !important;
                    border-right-width: 0px !important;
                    border-left-width: 0px !important;
                    border-top-width: 15px !important;
                    border-bottom-width: 15px !important
                }

                .es-adaptive table,
                .es-left,
                .es-right {
                    width: 100% !important
                }

                .es-content table,
                .es-header table,
                .es-footer table,
                .es-content,
                .es-footer,
                .es-header {
                    width: 100% !important;
                    max-width: 600px !important
                }

                .es-adapt-td {
                    display: block !important;
                    width: 100% !important
                }

                .adapt-img {
                    width: 100% !important;
                    height: auto !important
                }

                .es-m-p0 {
                    padding: 0px !important
                }

                .es-m-p0r {
                    padding-right: 0px !important
                }

                .es-m-p0l {
                    padding-left: 0px !important
                }

                .es-m-p0t {
                    padding-top: 0px !important
                }

                .es-m-p0b {
                    padding-bottom: 0 !important
                }

                .es-m-p20b {
                    padding-bottom: 20px !important
                }

                .es-mobile-hidden,
                .es-hidden {
                    display: none !important
                }

                tr.es-desk-hidden,
                td.es-desk-hidden,
                table.es-desk-hidden {
                    width: auto !important;
                    overflow: visible !important;
                    float: none !important;
                    max-height: inherit !important;
                    line-height: inherit !important
                }

                tr.es-desk-hidden {
                    display: table-row !important
                }

                table.es-desk-hidden {
                    display: table !important
                }

                td.es-desk-menu-hidden {
                    display: table-cell !important
                }

                .es-menu td {
                    width: 1% !important
                }

                table.es-table-not-adapt,
                .esd-block-html table {
                    width: auto !important
                }

                table.es-social {
                    display: inline-block !important
                }

                table.es-social td {
                    display: inline-block !important
                }

                .es-desk-hidden {
                    display: table-row !important;
                    width: auto !important;
                    overflow: visible !important;
                    max-height: inherit !important
                }
            }

            @media screen and (max-width:384px) {
                .mail-message-content {
                    width: 414px !important
                }
            }
        </style>
    </head>

    <body
        style="width:100%;font-family:arial, 'helvetica neue', helvetica, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
        <div dir="ltr" class="es-wrapper-color" lang="en" style="background-color:#ffffff">
            <!--[if gte mso 9]><v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t"> <v:fill type="tile" color="#ffffff"></v:fill> </v:background><![endif]-->
            <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" role="none"
                style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FFFFFF">
                <tr>
                    <td valign="top" style="padding:0;Margin:0">
                        <table cellpadding="0" cellspacing="0" class="es-footer" align="center" role="none"
                            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
                            <tr>
                                <td align="center" bgcolor="#000000" style="padding:0;Margin:0;background-color:#000000">
                                    <table bgcolor="#bcb8b1" class="es-footer-body" align="center" cellpadding="0"
                                        cellspacing="0" role="none"
                                        style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
                                        <tr>
                                            <td align="left" bgcolor="#070606"
                                                style="Margin:0;padding-top:20px;padding-bottom:20px;padding-left:40px;padding-right:40px;background-color:#070606">
                                                <table cellpadding="0" cellspacing="0" width="100%" role="none"
                                                    style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                    <tr>
                                                        <td align="center" valign="top"
                                                            style="padding:0;Margin:0;width:520px">
                                                            <table cellpadding="0" cellspacing="0" width="100%"
                                                                role="presentation"
                                                                style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                                <tr>
                                                                    <td align="center"
                                                                        style="padding:0;Margin:0;font-size:0px"><a
                                                                            target="_blank"
                                                                            href="https://asthra2k24.netlify.app"
                                                                            style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#2D3142;font-size:14px"><img
                                                                                src="https://i.ibb.co/qRVvvQc/ASTHRA-Logo.png"
                                                                                alt="Logo"
                                                                                style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic"
                                                                                height="150" title="Logo" width="210"></a>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none"
                            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
                            <tr>
                                <td align="center" bgcolor="#000000" style="padding:0;Margin:0;background-color:#000000">
                                    <table bgcolor="#efefef" class="es-content-body" align="center" cellpadding="0"
                                        cellspacing="0"
                                        style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#EFEFEF;border-radius:20px 20px 0 0;width:600px"
                                        role="none">
                                        <tr>
                                            <td align="left"
                                                style="padding:0;Margin:0;padding-top:20px;padding-left:40px;padding-right:40px">
                                                <table cellpadding="0" cellspacing="0" width="100%" role="none"
                                                    style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                    <tr>
                                                        <td align="center" valign="top"
                                                            style="padding:0;Margin:0;width:520px">
                                                            <table cellpadding="0" cellspacing="0" width="100%"
                                                                bgcolor="#fafafa"
                                                                style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:separate;border-spacing:0px;background-color:#fafafa;border-radius:10px"
                                                                role="presentation">
                                                                <tr>
                                                                    <td align="left" style="padding:20px;Margin:0">
                                                                        <h3
                                                                            style="Margin:0;line-height:34px;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;font-size:28px;font-style:normal;font-weight:bold;color:#2D3142">
                                                                            Dear ${participantName},</h3>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <br></p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            We are thrilled to welcome you to ASTHRA 2K24!
                                                                        </p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <br></p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            Your registration for ASTHRA 2K24 has been
                                                                            successfully recorded. Our team is currently
                                                                            processing your verification, which may take a
                                                                            little time. Rest assured, we'll notify you as
                                                                            soon as it's completed.</p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <br></p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            Thank you for registering for ASTHRA 2K24. If
                                                                            you have any queries, please feel free to reach
                                                                            out to our team.</p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <br></p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            Contact the event coordinator, <strong>Mr. S.
                                                                                Sethupathi</strong>, at Phone no:
                                                                            <strong>+916380295331</strong>.</p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <br></p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            Thanks,</p>
                                                                        <p
                                                                            style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:Imprima, Arial, sans-serif;line-height:27px;color:#2D3142;font-size:18px">
                                                                            <strong><a
                                                                                    href="https://www.linkedin.com/company/innak/"
                                                                                    target="_blank"
                                                                                    style="-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;text-decoration:underline;color:#2D3142;font-size:18px">Asthra
                                                                                    Tech Team</a></strong>
                                                                                    
                                                                                    </p>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none"
                        style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
                        <tr>
                            <td align="center" bgcolor="#000" style="padding:0;Margin:0;background-color:#000">
                                <table bgcolor="#efefef" class="es-content-body" align="center" cellpadding="0"
                                    cellspacing="0"
                                    style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#EFEFEF;border-radius:0 0 20px 20px;width:600px"
                                    role="none">
                                    <tr>
                                        <td class="esdev-adapt-off" align="left"
                                            style="padding:0;Margin:0;padding-left:40px;padding-right:40px">
                                            <table cellpadding="0" cellspacing="0" width="100%" role="none"
                                                style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                <tr>
                                                    <td align="center" valign="top"
                                                        style="padding:0;Margin:0;width:520px">
                                                        <table cellpadding="0" cellspacing="0" width="100%"
                                                            role="presentation"
                                                            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                            <tr>
                                                                <td align="center"
                                                                    style="padding:0;Margin:0;padding-bottom:10px;padding-top:20px;font-size:0">
                                                                   
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                        <table cellpadding="0" cellspacing="0" class="es-footer" align="center" role="none"
                            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
                            <tr>
                                <td align="center" bgcolor="#000000" style="padding:0;Margin:0;background-color:#000000">
                                    <table bgcolor="#bcb8b1" class="es-footer-body" align="center" cellpadding="0"
                                        cellspacing="0" role="none"
                                        style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
                                        <tr>
                                            <td align="left" bgcolor="#070606"
                                                style="Margin:0;padding-top:20px;padding-bottom:20px;padding-left:40px;padding-right:40px;background-color:#070606">
                                                <table cellpadding="0" cellspacing="0" width="100%" role="none"
                                                    style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                    <tr>
                                                        <td align="center" valign="top"
                                                            style="padding:0;Margin:0;width:520px">
                                                            <table cellpadding="0" cellspacing="0" width="100%"
                                                                role="presentation"
                                                                style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                                <tr>
                                                                    <td align="center" style="padding: 0; Margin: 0; font-size: 0px;">
                                                                        <h2 style="margin: 0; padding: 0; font-size: 18px; color: #ffffff; font-weight: bold; margin-bottom: 20px;">Tech Support</h2>
                                                                    
                                                                        <a target="_blank" href="https://www.instagram.com/innak_official/" style="-webkit-text-size-adjust: none; -ms-text-size-adjust: none; mso-line-height-rule: exactly; text-decoration: underline; color: #2D3142; font-size: 14px; display: block;">
                                                                    
                                                                            
                                                                            <img src="https://ik.imagekit.io/dayanidi/INNAK.png" alt="Logo" style="display: block; border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic;" height="" title="Logo" width="210">
                                                                        </a>
                                                                    </td>
                                                                    
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
    </body>

    </html>`;

  try {
    await MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: formattedBody
    });
    row.push("Email sent successfully!");
    row.push("");
  } catch (error) {
    row.push("Email sent successfully!");
    row.push(error);
    // Handle the error as needed (e.g., notify the user, log additional details, etc.)
  }

  return row;

}
