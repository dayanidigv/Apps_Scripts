function sendRegistrationConfirmationEmail(recipient) {
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
        <title>New Template 2</title>
        <!--[if (mso 16)]><style type="text/css">     a {text-decoration: none;}     </style><![endif]-->
        <!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--> <!--[if gte mso 9]><xml> <o:OfficeDocumentSettings> <o:AllowPNG></o:AllowPNG> <o:PixelsPerInch>96</o:PixelsPerInch> </o:OfficeDocumentSettings> </xml>
        <![endif]--> <!--[if !mso]><!-- -->
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
        <div dir="ltr" class="es-wrapper-color" lang="en" style="background-color:#FFFFFF">
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
                                                                            Dear Participant,</h3>
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
                                                                                    Tech Team</a></strong></p>
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
                                                                        style="padding:0;Margin:0;padding-bottom:20px;padding-top:40px;font-size:0">
                                                                        <table border="0" width="100%" height="100%"
                                                                            cellpadding="0" cellspacing="0"
                                                                            role="presentation"
                                                                            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                                                            <tr>
                                                                                <td
                                                                                    style="padding:0;Margin:0;border-bottom:1px solid #666666;background:unset;height:1px;width:100%;margin:0px">
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
                    </td>
                </tr>
            </table>
        </div>
    </body>

    </html>`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: formattedBody
  });
}

function sendMultipleMail(){
  sendRegistrationConfirmationEmail("<gmail>");
}
