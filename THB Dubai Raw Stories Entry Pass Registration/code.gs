function getData() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();

  var isSheet = sheets.getSheetByName("Main");
  if (isSheet) {
    var data = isSheet.getDataRange().getValues();

    data.forEach(function (entry) {
      Logger.log(entry);
    });
  }

}

function generateID() {
  var date = new Date();
  var year = date.getFullYear().toString().substr(-2);
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);
  var hours = ("0" + date.getHours()).slice(-2);
  var minutes = ("0" + date.getMinutes()).slice(-2);
  var seconds = ("0" + date.getSeconds()).slice(-2);
  var milliseconds = ("00" + date.getMilliseconds()).slice(-3);

  // Combine date components to form ID code
  var idCode = year + month + day + hours + minutes + seconds + milliseconds;

  return idCode
}

function doPost(e) {
  try {
    var formData = JSON.parse(e.postData.contents);

    // Validate form data
    if (!formData || !formData.category || !formData.name || !formData.email || !formData.mobile || !formData.gender || !formData.age_group || !formData.no_of_pass) {
      throw new Error(`Invalid form data. Please fill in all required fields.`);
    }

    // Ensure that formData.no_of_pass is a valid number
    if (isNaN(parseInt(formData.no_of_pass))) {
      throw new Error("Invalid number of passes. Please enter a valid number.");
    }

    // Ensure that formData.category is a valid category
    if (formData.category !== "speaker" && formData.category !== "partner" && formData.category !== "tepa" && formData.category !== "normal") {
      throw new Error("Invalid Category. Please make sure you select a valid category.");
    }


    if (formData.isUpdate) {
      handleFormUpdate(formData);
    } else {

      handleFormSubmission(formData);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'message': "Registration details received successfully." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error processing form submission:", error);
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'message': error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

}

async function handleFormSubmission(formData) {
  try {
    var sheets = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = sheets.getSheetByName("Main");

    var date = new Date();
    var iD = generateID();

    var row = [
      date,
      iD,
      formData.category,
      formData.name,
      formData.email,
      formData.mobile,
      formData.gender,
      formData.age_group,
      formData.linkedin,
      formData.facebook,
      formData.instagram,
      parseInt(formData.no_of_pass) + 1,
    ];

    mainSheet.appendRow(row);


    if (parseInt(formData.no_of_pass) !== 0) {
      for (var i = 1; i <= parseInt(formData.no_of_pass); i++) {
        mainSheet.appendRow(["", iD, `${formData.category} Guest`, formData[`name_${i}`], formData[`email_${i}`], formData[`mobile_${i}`], formData[`gender_${i}`], formData[`age_group_${i}`]]);
      }
    }


    // Determine the registration sheet based on the category
    var registrationSheetName = formData.category + " registration";
    var registrationSheet = sheets.getSheetByName(registrationSheetName);

    if (registrationSheet) {
      var registrationRow;

      if (formData.category == "partner") {
        registrationRow = [
          date,
          iD,
          formData.partner_category,
          formData.name,
          formData.companyOrBusinessName,
          formData.industryType,
          formData.designationOrRole,
          formData.email,
          formData.mobile,
          formData.gender,
          formData.age_group,
          formData.linkedin,
          formData.facebook,
          formData.instagram,
          parseInt(formData.no_of_pass) + 1,
        ];
      } else {
        registrationRow = [
          date,
          iD,
          formData.name,
          formData.email,
          formData.mobile,
          formData.gender,
          formData.age_group,
          formData.linkedin,
          formData.facebook,
          formData.instagram,
          parseInt(formData.no_of_pass) + 1,
        ];
      }
      registrationSheet.appendRow(registrationRow);
      sendRegistrationSuccessEmail(formData, iD) ;
    } else {
      console.error("Registration sheet not found for category:", formData.category);
      throw new Error("Error handling form submission.");
    }

  } catch (error) {
    console.error("Error handling form submission:", error);
    throw new Error("Error handling form submission.");
  }
}

async function handleFormUpdate(formData) {
  return true;
}


function doGet(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var isInputSheet = e.parameter.getInputByName;
  var isverify = e.parameter.isverify;
  var verifyKey = e.parameter.verifyKey;
  var verifyValue = e.parameter.verifyValue;
  var verifyCategory = e.parameter.verifyCategory;
  var verifyToGet = e.parameter.verifyToGet;
  var getDataCountByName = e.parameter.getDataCountByName;

  if (isInputSheet) {
    var sheet;
    if (isInputSheet == "spn") {
      sheet = spreadsheet.getSheetByName("Speakers Name");
    } else if (isInputSheet == "pc") {
      sheet = spreadsheet.getSheetByName("Partner  Category");
    } else {
      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Name is not Match' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    var jsonData = convertToJson(values);

    return ContentService.createTextOutput(JSON.stringify(jsonData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  else if (isverify == "true") {
    // Validate form data
    if (!verifyKey || !verifyValue || !verifyCategory) {
      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Incomplete verification parameters. Please make sure all verification parameters are provided.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (verifyCategory == "speaker") {
      const sheet = spreadsheet.getSheetByName("speaker registration");
      var data = sheet.getDataRange().getValues();
      var isAvailable = false;
      if (verifyKey == "name") {
        for (var i = 0; i < data.length; i++) {
          if (data[i][2] === verifyValue) {
            isAvailable = true;
            if (verifyToGet) {
              if (verifyToGet == "id") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][1] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "email") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][3] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": "" }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
            return ContentService
              .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available` }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      } else if (verifyKey == "email") {
        for (var i = 0; i < data.length; i++) {
          if (data[i][3] === verifyValue) {
            isAvailable = true;
            if (verifyToGet) {
              if (verifyToGet == "id") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][1] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "all") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": "" }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
            return ContentService
              .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "iD": data[i][1] }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      if (!isAvailable) {
        return ContentService
          .createTextOutput(JSON.stringify({ 'result': 'error', 'message': `${verifyValue} is not available` }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    else if (verifyCategory == "partner") {
      const sheet = spreadsheet.getSheetByName("partner registration");
      var data = sheet.getDataRange().getValues();
      var isAvailable = false;
      if (verifyKey == "email") {
        for (var i = 0; i < data.length; i++) {
          if (data[i][7] === verifyValue) {
            isAvailable = true;
            if (verifyToGet) {
              if (verifyToGet == "id") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][1] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "all") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": "" }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
            return ContentService
              .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "iD": data[i][1] }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
        if (!isAvailable) {
          return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'error', 'message': `${verifyValue} is not available` }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } else if (verifyKey == "category") {
        for (var i = 0; i < data.length; i++) {
          if (data[i][2] === verifyValue) {
            isAvailable = true;
            if (verifyToGet) {
              if (verifyToGet == "id") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][1] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "all") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "email") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][7] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": "" }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
            return ContentService
              .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "iD": data[i][1] }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
        if (!isAvailable) {
          return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'error', 'message': `${verifyValue} is not available` }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      else {
        return ContentService
          .createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'verifykey value is not matched' }))
          .setMimeType(ContentService.MimeType.JSON);

      }
    }
    else if (verifyCategory == "tepa") {
      const sheet = spreadsheet.getSheetByName("tepa registration");
      var data = sheet.getDataRange().getValues();
      var isAvailable = false;
      if (verifyKey == "email") {
        for (var i = 0; i < data.length; i++) {
          if (data[i][3] === verifyValue) {
            isAvailable = true;
            if (verifyToGet) {
              if (verifyToGet == "id") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i][1] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else if (verifyToGet == "all") {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": data[i] }))
                  .setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService
                  .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "returnData": "" }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
            return ContentService
              .createTextOutput(JSON.stringify({ 'result': 'success', 'message': `${verifyValue} is available`, "iD": data[i][1] }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
        if (!isAvailable) {
          return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'error', 'message': `${verifyValue} is not available` }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

    }
  }

  else if (getDataCountByName) {

    if (getDataCountByName !== "all" && getDataCountByName !== "speaker" && getDataCountByName !== "partner" && getDataCountByName !== "tepa") {
      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Invalid data count name.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (getDataCountByName === "tepa") {
      const sheet = spreadsheet.getSheetByName("tepa registration");
      var data = sheet.getDataRange().getValues();
      var dataCount = data.length;
      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'success', 'dataCount': dataCount - 1 }))
        .setMimeType(ContentService.MimeType.JSON);
    }


  }



  return ContentService
    .createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Unauthorized access' }))
    .setMimeType(ContentService.MimeType.JSON);

}

function convertToJson(values) {
  var headers = values[0];
  var jsonData = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};

    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }

    jsonData.push(obj);
  }

  return jsonData;
}

function sendMail(){
  sendRegistrationSuccessEmail({category: "Tepa",no_of_pass : "0",name : "daya", email : "dayanidigv954@gmail.com"},"13356");
}

async function sendRegistrationSuccessEmail(formData, iD) {
  var subject = "Confirmation: Entry Pass for Dubai Raw Stories on May 26, 2024";

  const _category = formData.category;
  const pass_Category = _category === "speaker" ? "Speaker"
                      : _category === "partner" ? "Sponsor"
                      : "Tepa";


  var number_of_free_passes = parseInt(formData.no_of_pass) + 1 >= 2 ? 2 : 1;
  var number_of_paid_passes = parseInt(formData.no_of_pass) + 1 >= 2 ? parseInt(formData.no_of_pass) - 1 : 0;
  var number_of_paid_passes_text = number_of_paid_passes === 0 ?  "" 
                                : `<p><strong>Number of Paid Passes: </strong> ${number_of_paid_passes} (Payment to be made on-site)</p>`;

  let tableData = `<tr>
                  <td style="text-align: center;">${formData.name}</td>
                  <td style="text-align: center;">${formData.email}</td>
                </tr>`;

  if (parseInt(formData.no_of_pass) !== 0) {
    for (var i = 1; i <= parseInt(formData.no_of_pass); i++) {
      tableData += `
      <tr>
          <td style="text-align: center;">${formData[`name_${i}`]}</td>
          <td style="text-align: center;">${formData[`email_${i}`]}</td>
      </tr>`;
    }
  }



  var formattedBody = `<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Entry Pass Confirmation</title>
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,100..900;1,100..900&display=swap');
            

            body {
                font-family: "Montserrat", sans-serif;
                font-size: 16px;
                margin: 0;
                padding: 0;
            }

            .highlight {
                color: #C41804;
            }

            .container {
                max-width: 680px;
                margin: 0 auto;
            }

            img {
                max-width: 100%;
                height: 100px;
            }

            p {
                margin: 5px 0;
                line-height: 1.4;
            }

            .icon {
                width: 23px;
                height: 23px;
                border: #C41804 solid 2px;
                border-radius: 50px;
                display: inline-block;
                justify-content: center;
                text-align: center;
                color: #C41804;
                margin-right: 5px;
            }

            .icon i {
                font-weight: bolder;
                margin-top: 5px;
            }

            table {
                width: 100%;
                border-collapse: collapse;
            }

            th,
            td {
                padding: 8px;
                border-bottom: 1px solid #ddd;
                text-align: left;
            }

            th {
                background-color: #f2f2f2;
            }

            @media only screen and (max-width: 600px) {
                .container {
                    width: 100% !important;
                }

              
            }
        </style>
    </head>

    <body style="margin: 0; padding: 0;">

        <!-- Header with Logos -->
        <table class="container" style="width: 100%; border-collapse: collapse;">
            <tr>
                <td style="padding: 20px; text-align: center;">
                    <img src="https://dayanidigv.github.io/THB/thb-raw-stories.png" alt="thehalfbrick Raw Stories"
                        style="max-width: 100%; height: 100px;" />
                </td>
            </tr>
        </table>

        <!-- Confirmation Message -->
        <table class="container"
            style="width: 100%; border-collapse: collapse; border-top: 5px solid #C41804; border-bottom: 5px solid #C41804; background-image: url('https://thehalfbrick.com/wp-content/uploads/2022/08/190085d2-a19d-4b44-a2ac-b56ade048c24-1536x981.png'); background-repeat: no-repeat; background-position: center center;">
            <tr>
                <td style="padding: 20px; background-color: #f8f8f8;">
                    <h1 style="margin-top: 0;" class="highlight">Confirmation of Entry Pass</h1>
                    <p>Your entry pass has been successfully confirmed. Kindly review the details below:</p>
                    <p><strong>Event:</strong> Dubai raw stories</p>
                    <p><strong>Date:</strong> 26 May, 2024</p>
                    <p><strong>Location:</strong> Hotel Sheraton Grand, Starlight Ballroom, 3 Sheikh Zayed Rd, P.O. Box
                        123979, Dubai, UAE</p><br>

                    <p><strong>Pass ID:</strong> ${iD}</p><br>
                    <table style="margin: auto; max-width:  500px;">
                        <tr>
                            <th style="text-align: center;">Name</th>
                            <th style="text-align: center;">Email</th>
                        </tr>
                        ${tableData}
                    </table>

                    <br>
                    <p><strong>Pass Category:</strong> ${pass_Category}</p><br>
                    <p><strong>Number of Free Passes:</strong> ${number_of_free_passes} </p>
                    ${number_of_paid_passes_text}<br>
                    <p>Thank you for choosing to attend Dubai Raw Stories.<br> We look forward to welcoming you at the
                        event!</p>
                </td>
            </tr>
        </table>

    </body>

    </html>`;

  try {
    await MailApp.sendEmail({
      to: formData.email,
      subject: subject,
      htmlBody: formattedBody
    });
    return iD; 
  } catch (error) {
    Logger.log(error);
    throw new Error("Failed to send email. Please try again later."); 
    // Handle the error as needed (e.g., notify the user, log additional details, etc.)
  }

}
