const APP_CONFIG = Object.freeze({
  agencyName: 'Bill Layne Insurance Agency',
  agencyPhone: '336-835-1993',
  agencyEmail: 'docs@billlayneinsurance.com',
  agencyWebsite: 'https://mynolossform.com',
  defaultTimeZone: 'America/New_York',
  defaultPortalSecret: 'BillLayneInsurance2025',
  rootFolderName: 'Bill Layne Insurance - No Loss Statements'
});

function doGet(e) {
  const payload = {
    ok: true,
    app: 'no-loss-statement',
    version: '2026-04-01',
    timestamp: new Date().toISOString()
  };

  const callback = e && e.parameter ? e.parameter.callback : '';
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse_(payload);
}

function doPost(e) {
  try {
    const payload = parseRequestBody_(e);

    if (payload && payload.action === 'send_link_sms') {
      return handleAgentPortalSms_(payload);
    }

    return handleStatementSubmission_(payload);
  } catch (error) {
    console.error('No Loss backend failure: ' + (error && error.stack ? error.stack : error));
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function parseRequestBody_(e) {
  const contents = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!contents) {
    throw new Error('Missing POST body.');
  }

  try {
    return JSON.parse(contents);
  } catch (error) {
    throw new Error('Could not parse JSON payload.');
  }
}

function handleStatementSubmission_(payload) {
  const runtime = getRuntimeConfig_();
  const submission = normalizeSubmission_(payload);
  const folder = createSubmissionFolder_(runtime.rootFolder, submission);
  const signatureBlob = createSignatureBlob_(submission.signatureUrl, submission.confirmationNumber);
  const signatureFile = folder.createFile(signatureBlob.copyBlob().setName(submission.confirmationNumber + ' - signature.png'));
  const archiveFile = folder.createFile(
    submission.confirmationNumber + ' - submission.json',
    JSON.stringify(buildArchiveObject_(submission), null, 2),
    MimeType.PLAIN_TEXT
  );
  const pdfFile = createStatementPdf_(folder, submission, signatureBlob);

  const officeEmailSent = sendOfficeEmail_(runtime, submission, folder, pdfFile, archiveFile, signatureFile);
  const agentCopySent = tryOptionalNotification_(function() {
    return sendAgentCopy_(runtime, submission, folder, pdfFile);
  }, 'agent copy email');
  const customerEmailSent = tryOptionalNotification_(function() {
    return sendCustomerEmail_(runtime, submission, pdfFile);
  }, 'customer confirmation email');
  const customerSmsSent = tryOptionalNotification_(function() {
    return sendCustomerConfirmationSms_(runtime, submission);
  }, 'customer confirmation sms');

  return jsonResponse_({
    ok: true,
    confirmationNumber: submission.confirmationNumber,
    driveFolderUrl: folder.getUrl(),
    pdfUrl: pdfFile.getUrl(),
    officeEmailSent: officeEmailSent,
    agentCopySent: agentCopySent,
    customerEmailSent: customerEmailSent,
    customerSmsSent: customerSmsSent
  });
}

function handleAgentPortalSms_(payload) {
  validateAgentPortalToken_(cleanText_(payload.authToken));

  const runtime = getRuntimeConfig_();
  const to = normalizePhone_(payload.to);
  const message = appendOptOut_(cleanText_(payload.message));

  if (!to) {
    throw new Error('Missing phone number.');
  }
  if (!message) {
    throw new Error('Missing SMS message.');
  }

  const smsResult = sendSmsViaTwilio_(runtime, to, message);

  return jsonResponse_({
    ok: true,
    action: 'send_link_sms',
    sid: smsResult.sid || '',
    status: smsResult.status || 'queued'
  });
}

function normalizeSubmission_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload must be an object.');
  }

  const submission = {
    confirmationNumber: cleanText_(payload.confirmationNumber) || generateFallbackConfirmation_(),
    sessionId: cleanText_(payload.sessionId),
    insuredName: requireField_(payload.insuredName, 'insuredName'),
    email: cleanText_(payload.email),
    phone: requireField_(payload.phone, 'phone'),
    propertyAddress: requireField_(payload.propertyAddress, 'propertyAddress'),
    city: requireField_(payload.city, 'city'),
    state: requireField_(payload.state, 'state'),
    zipCode: requireField_(payload.zipCode, 'zipCode'),
    insuranceCompany: requireField_(payload.insuranceCompany, 'insuranceCompany'),
    policyNumber: requireField_(payload.policyNumber, 'policyNumber'),
    policyType: requireField_(payload.policyType, 'policyType'),
    amountPaid: cleanText_(payload.amountPaid),
    cancellationDate: requireField_(payload.cancellationDate, 'cancellationDate'),
    reinstatementDate: requireField_(payload.reinstatementDate, 'reinstatementDate'),
    noLossConfirmation: truthyToYesNo_(payload.noLossConfirmation, true),
    dmvAcknowledgement: truthyToYesNo_(payload.dmvAcknowledgement),
    mortgageAcknowledgement: truthyToYesNo_(payload.mortgageAcknowledgement),
    agencyName: cleanText_(payload.agencyName) || APP_CONFIG.agencyName,
    agencyAddress: cleanText_(payload.agencyAddress),
    agencyEmail: cleanText_(payload.agencyEmail) || APP_CONFIG.agencyEmail,
    agencyPhone: cleanText_(payload.agencyPhone) || APP_CONFIG.agencyPhone,
    agentName: cleanText_(payload.agentName),
    agentEmail: cleanText_(payload.agentEmail),
    signatureUrl: cleanText_(payload.signatureUrl || payload.signature),
    signatureDateTime: cleanText_(payload.signatureDateTime),
    signatureTime: cleanText_(payload.signatureTime),
    submittedAt: cleanText_(payload.submittedAt) || new Date().toISOString(),
    timezone: cleanText_(payload.timezone) || APP_CONFIG.defaultTimeZone,
    ipAddress: cleanText_(payload.ipAddress),
    browserFingerprint: cleanText_(payload.browserFingerprint),
    deviceInfo: cleanText_(payload.deviceInfo),
    screenResolution: cleanText_(payload.screenResolution),
    userAgent: cleanText_(payload.userAgent),
    formVersion: cleanText_(payload.formVersion),
    submissionMethod: cleanText_(payload.submissionMethod),
    signatureMetadata: parseMaybeJson_(payload.signatureMetadata)
  };

  if (!submission.signatureUrl || submission.signatureUrl.indexOf('data:image') !== 0) {
    throw new Error('Missing signature image.');
  }

  return submission;
}

function getRuntimeConfig_() {
  const props = PropertiesService.getScriptProperties();
  const rootFolder = ensureRootFolder_(props);
  const officeEmails = String(props.getProperty('OFFICE_EMAILS') || APP_CONFIG.agencyEmail)
    .split(',')
    .map(function(value) { return cleanText_(value); })
    .filter(Boolean);

  return {
    rootFolder: rootFolder,
    officeEmails: officeEmails,
    fromName: cleanText_(props.getProperty('FROM_NAME')) || APP_CONFIG.agencyName,
    customerReplyTo: cleanText_(props.getProperty('CUSTOMER_REPLY_TO')) || officeEmails[0] || APP_CONFIG.agencyEmail,
    timeZone: cleanText_(props.getProperty('TIMEZONE')) || Session.getScriptTimeZone() || APP_CONFIG.defaultTimeZone,
    portalSecret: cleanText_(props.getProperty('AGENT_PORTAL_SECRET')) || APP_CONFIG.defaultPortalSecret,
    twilioSid: cleanText_(props.getProperty('TWILIO_SID')),
    twilioToken: cleanText_(props.getProperty('TWILIO_TOKEN')),
    twilioFrom: cleanText_(props.getProperty('TWILIO_FROM')),
    twilioMessagingServiceSid: cleanText_(props.getProperty('TWILIO_MESSAGING_SERVICE_SID'))
  };
}

function ensureRootFolder_(props) {
  const existingId = cleanText_(props.getProperty('UPLOADS_ROOT_FOLDER_ID'));
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (error) {
      console.error('Configured UPLOADS_ROOT_FOLDER_ID was not accessible, creating a new root folder: ' + error);
    }
  }

  const folderName = cleanText_(props.getProperty('UPLOADS_ROOT_FOLDER_NAME')) || APP_CONFIG.rootFolderName;
  const folder = DriveApp.createFolder(folderName);
  props.setProperty('UPLOADS_ROOT_FOLDER_ID', folder.getId());
  props.setProperty('UPLOADS_ROOT_FOLDER_NAME', folderName);
  return folder;
}

function createSubmissionFolder_(rootFolder, submission) {
  const submittedAt = safeDate_(submission.submittedAt);
  const yearMonth = Utilities.formatDate(submittedAt, APP_CONFIG.defaultTimeZone, 'yyyy-MM');
  const day = Utilities.formatDate(submittedAt, APP_CONFIG.defaultTimeZone, 'yyyy-MM-dd');

  const monthFolder = findOrCreateFolder_(rootFolder, yearMonth);
  const dayFolder = findOrCreateFolder_(monthFolder, day);
  const folderName = sanitizeFileName_(submission.confirmationNumber + ' - ' + submission.insuredName);

  return dayFolder.createFolder(folderName);
}

function createSignatureBlob_(signatureUrl, confirmationNumber) {
  const parts = signatureUrl.split(',');
  if (parts.length < 2) {
    throw new Error('Invalid signature data URL.');
  }

  const header = parts[0];
  const base64Data = parts[1];
  const mimeMatch = header.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64$/);
  const mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
  const extension = mimeType.split('/')[1] || 'png';
  const bytes = Utilities.base64Decode(base64Data);

  return Utilities.newBlob(bytes, mimeType, confirmationNumber + ' - signature.' + extension);
}

function createStatementPdf_(folder, submission, signatureBlob) {
  var docName = submission.confirmationNumber + ' - Statement of No Loss';
  var doc = DocumentApp.create(docName);
  var docId = doc.getId();
  var body = doc.getBody();

  // Set margins for single-page letterhead
  body.setMarginTop(28);
  body.setMarginBottom(20);
  body.setMarginLeft(45);
  body.setMarginRight(45);

  var navy = '#1a2744';
  var darkRed = '#8b1a1a';
  var gray = '#555555';

  // ── AGENCY HEADER (centered) ──
  var h1 = body.appendParagraph('BILL LAYNE INSURANCE AGENCY');
  h1.setFontSize(16).setBold(true).setForegroundColor(navy).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(0).setSpacingBefore(0);

  var h2 = body.appendParagraph('1283 N Bridge St, Elkin, NC 28621 \u2022 (336) 835-1993 \u2022 Save@BillLayneInsurance.com');
  h2.setFontSize(8).setBold(false).setForegroundColor(gray).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2).setSpacingBefore(0);

  var h3 = body.appendParagraph('STATEMENT OF NO LOSS - CONF #' + submission.confirmationNumber);
  h3.setFontSize(12).setBold(true).setForegroundColor(navy).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2).setSpacingBefore(4);

  // Red divider line
  body.appendHorizontalRule();

  // ── POLICY INFO TABLE (two columns) ──
  var table = body.appendTable();
  table.setBorderWidth(0);

  var row1 = table.appendTableRow();
  var c1 = row1.appendTableCell('INSURED: ' + (submission.insuredName || ''));
  c1.setWidth(280);
  formatCell_(c1, 9, true);
  var c2 = row1.appendTableCell('POLICY: ' + (submission.policyNumber || ''));
  formatCell_(c2, 9, true);

  // Address row
  var row1b = table.appendTableRow();
  var addr = (submission.propertyAddress || '') + ', ' + (submission.city || '') + ', ' + (submission.state || '') + ', ' + (submission.zipCode || '');
  var c3 = row1b.appendTableCell(addr);
  formatCell_(c3, 8, false);
  var c4 = row1b.appendTableCell('');
  formatCell_(c4, 8, false);

  // Company / Type
  var row2 = table.appendTableRow();
  var c5 = row2.appendTableCell('COMPANY: ' + (submission.insuranceCompany || ''));
  formatCell_(c5, 9, false);
  var c6 = row2.appendTableCell('TYPE: ' + (submission.policyType || ''));
  formatCell_(c6, 9, false);

  // Lapse / Reinstate
  var row3 = table.appendTableRow();
  var c7 = row3.appendTableCell('LAPSE: ' + (submission.cancellationDate || ''));
  formatCell_(c7, 9, false);
  var c8 = row3.appendTableCell('REINSTATE: ' + (submission.reinstatementDate || ''));
  formatCell_(c8, 9, false);

  // Amount / Phone
  var row4 = table.appendTableRow();
  var c9 = row4.appendTableCell('Total Amount Required to Reinstate: ' + (submission.amountPaid || 'N/A'));
  formatCell_(c9, 9, true);
  var c10 = row4.appendTableCell('PHONE: ' + (submission.phone || ''));
  formatCell_(c10, 9, false);
  c10.editAsText().setItalic(true);

  // Fees note
  var feeNote = body.appendParagraph('*Fees included in Total Amount Required to Reinstate');
  feeNote.setFontSize(7).setItalic(true).setForegroundColor(gray).setSpacingAfter(4).setSpacingBefore(0);

  // ── FULL STATEMENT TEXT ──
  var stmtFull = 'STATEMENT OF NO LOSS: I, ' + (submission.insuredName || '') + ', state that neither I nor any other person covered by this policy has had a claim or loss or been involved in an accident since the cancellation or expiration of the policy (the "no loss period") wherein this policy, including any and all coverages endorsed upon or made part of the policy may apply. In addition, if this reinstatement is for a personal or commercial auto, motorcycle, or RV policy, I certify that I have disclosed the current garaging location and use of all insured vehicles including if any such vehicle is used to deliver food or goods, to transport people for compensation, or for any other business purpose. I have also disclosed household members who are age 14 or older, and all persons who regularly drive any vehicle insured under this policy. I understand that this insurance company is relying solely upon this Statement of No Loss all of which is material, as an inducement to reinstate my policy with no lapse in coverage. I further understand that if a claim, loss, or accident has occurred during the no loss period, or if I failed to disclose the current garaging location and primary use of all vehicles insured under this policy, all persons who regularly drive these vehicles, and all members of my household who are age 14 or older, the reinstatement is null and void, my policy remains cancelled and no insurance coverage shall be provided. I agree that if my check or other payment for this reinstatement is not honored for any reason, the reinstatement is null and void and no coverage shall exist under this policy. I agree to pay a reinstatement fee and late fee (if applicable) in addition to the premium required to reinstate my policy. My payment will be applied first to the reinstatement and late fee and the remainder to the premium.';

  var stmtPara = body.appendParagraph(stmtFull);
  stmtPara.setFontSize(7.5).setLineSpacing(1.0).setSpacingAfter(4).setSpacingBefore(4);
  // Bold the label
  stmtPara.editAsText().setBold(0, 21, true);

  // Warning
  var warning = body.appendParagraph('\u26A0 WARNING: It is a crime to knowingly provide false information to an insurance company. Penalties include imprisonment, fines, and denial of benefits.');
  warning.setFontSize(8).setForegroundColor(darkRed).setItalic(true).setSpacingAfter(2).setSpacingBefore(2);

  // Checkbox acknowledgement
  var ack = body.appendParagraph('\u2611 I agree to all terms above and confirm NO LOSS occurred during the lapse period.');
  ack.setFontSize(9).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2).setSpacingBefore(2);

  // Red divider
  body.appendHorizontalRule();

  // ── ELECTRONIC SIGNATURE ──
  var sigTitle = body.appendParagraph('ELECTRONIC SIGNATURE');
  sigTitle.setFontSize(9).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2).setSpacingBefore(2);

  // Signature image
  var sigImg = body.appendImage(signatureBlob);
  sigImg.setWidth(160).setHeight(60);
  var sigImgPara = sigImg.getParent().asParagraph();
  sigImgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Signature metadata line
  var sigDateTime = submission.signatureDateTime || formatDateTime_(submission.submittedAt);
  var sigMeta = body.appendParagraph((submission.insuredName || '') + ' \u2022 ' + sigDateTime + ' \u2022 IP: ' + (submission.ipAddress || 'N/A') + ' \u2022 Session: ' + (submission.sessionId || submission.confirmationNumber));
  sigMeta.setFontSize(6).setForegroundColor(gray).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2).setSpacingBefore(2);

  // Legal footer
  var legalLine = body.appendParagraph('Electronic signature valid per E-SIGN Act \u2022 Bill Layne Insurance Agency \u2022 www.BillLayneInsurance.com');
  legalLine.setFontSize(7).setForegroundColor(gray).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(0).setSpacingBefore(0);

  var feeLine = body.appendParagraph('* Fees included in Total Amount Required to Reinstate');
  feeLine.setFontSize(7).setItalic(true).setForegroundColor(gray).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(0).setSpacingBefore(0);

  doc.saveAndClose();

  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs(MimeType.PDF).setName(docName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);

  return pdfFile;
}

function formatCell_(cell, fontSize, bold) {
  cell.setBackgroundColor(null);
  var para = cell.getChild(0).asParagraph();
  para.setFontSize(fontSize).setSpacingAfter(1).setSpacingBefore(1);
  if (bold) para.setBold(true);
}

function sendOfficeEmail_(runtime, submission, folder, pdfFile, archiveFile, signatureFile) {
  const recipients = runtime.officeEmails.filter(Boolean);
  if (!recipients.length) {
    return false;
  }

  const subject = '[No Loss] ' + submission.insuredName + ' - ' + submission.policyNumber + ' - ' + submission.confirmationNumber;
  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;color:#1f2937;line-height:1.6;">',
    '<h2 style="margin:0 0 12px;">New Statement of No Loss Submission</h2>',
    '<p><strong>Confirmation #:</strong> ' + htmlEscape_(submission.confirmationNumber) + '<br>',
    '<strong>Insured:</strong> ' + htmlEscape_(submission.insuredName) + '<br>',
    '<strong>Policy Number:</strong> ' + htmlEscape_(submission.policyNumber) + '<br>',
    '<strong>Company:</strong> ' + htmlEscape_(submission.insuranceCompany) + '<br>',
    '<strong>Policy Type:</strong> ' + htmlEscape_(submission.policyType) + '<br>',
    '<strong>Email:</strong> ' + htmlEscape_(submission.email || 'Not provided') + '<br>',
    '<strong>Phone:</strong> ' + htmlEscape_(submission.phone) + '<br>',
    '<strong>Submitted:</strong> ' + htmlEscape_(formatDateTime_(submission.submittedAt)) + '</p>',
    '<p><strong>Drive Folder:</strong> <a href="' + htmlEscape_(folder.getUrl()) + '">' + htmlEscape_(folder.getUrl()) + '</a><br>',
    '<strong>PDF:</strong> <a href="' + htmlEscape_(pdfFile.getUrl()) + '">' + htmlEscape_(pdfFile.getName()) + '</a><br>',
    '<strong>JSON Archive:</strong> <a href="' + htmlEscape_(archiveFile.getUrl()) + '">' + htmlEscape_(archiveFile.getName()) + '</a><br>',
    '<strong>Signature:</strong> <a href="' + htmlEscape_(signatureFile.getUrl()) + '">' + htmlEscape_(signatureFile.getName()) + '</a></p>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: recipients[0],
    cc: recipients.slice(1).join(','),
    subject: subject,
    name: runtime.fromName,
    replyTo: isValidEmail_(submission.email) ? submission.email : runtime.customerReplyTo,
    htmlBody: htmlBody,
    body: stripHtml_(htmlBody),
    attachments: [
      pdfFile.getBlob().setName(pdfFile.getName()),
      archiveFile.getBlob().setName(archiveFile.getName())
    ]
  });

  return true;
}

function sendAgentCopy_(runtime, submission, folder, pdfFile) {
  if (!isValidEmail_(submission.agentEmail)) {
    return false;
  }

  const subject = 'Signed Statement of No Loss - ' + submission.insuredName + ' (' + submission.policyNumber + ')';
  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;color:#1f2937;line-height:1.6;">',
    '<p>The customer has completed their Statement of No Loss.</p>',
    '<p><strong>Confirmation #:</strong> ' + htmlEscape_(submission.confirmationNumber) + '<br>',
    '<strong>Insured:</strong> ' + htmlEscape_(submission.insuredName) + '<br>',
    '<strong>Policy Number:</strong> ' + htmlEscape_(submission.policyNumber) + '<br>',
    '<strong>Drive Folder:</strong> <a href="' + htmlEscape_(folder.getUrl()) + '">' + htmlEscape_(folder.getUrl()) + '</a></p>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: submission.agentEmail,
    subject: subject,
    name: runtime.fromName,
    replyTo: runtime.customerReplyTo,
    htmlBody: htmlBody,
    body: stripHtml_(htmlBody),
    attachments: [pdfFile.getBlob().setName(pdfFile.getName())]
  });

  return true;
}

function sendCustomerEmail_(runtime, submission, pdfFile) {
  if (!isValidEmail_(submission.email)) {
    return false;
  }

  var firstName = (submission.insuredName || 'there').split(' ')[0];
  var carrier = submission.insuranceCompany || 'Your Insurance Company';
  var localTime = '';
  try {
    localTime = Utilities.formatDate(new Date(submission.submittedAt), APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch(e) {
    localTime = submission.submittedAt || new Date().toLocaleString();
  }

  var subject = 'Statement of No Loss Received - ' + (submission.insuredName || '') + ' - Bill Layne Insurance';
  var logoUrl = 'https://i.imgur.com/lxu9nfT.png';
  var ff = "font-family:'Inter',Arial,'Helvetica Neue',Helvetica,sans-serif;";

  var htmlBody = [
    '<!DOCTYPE html><html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="x-apple-disable-message-reformatting"><meta name="format-detection" content="telephone=no"><title>Statement of No Loss Received - ' + htmlEscape_(submission.insuredName || '') + ' - Bill Layne Insurance</title>',
    '<!--[if mso]><noscript><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript><![endif]-->',
    '<style>body,table,td,p,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0pt;mso-table-rspace:0pt}img{-ms-interpolation-mode:bicubic;border:0;outline:none;text-decoration:none;display:block}body{margin:0!important;padding:0!important;background-color:#f1f5f9;width:100%!important}.card-pad{padding:28px 32px!important}.hero-pad{padding:36px 28px!important}@media only screen and (max-width:620px){.email-container{width:100%!important}.card-pad{padding:20px 16px!important}.hero-pad{padding:28px 16px!important}.cta-btn{width:100%!important}.btn-td{padding:14px 24px!important}}</style>',
    '<script type="application/ld+json">{"@context":"http://schema.org","@type":"EmailMessage","description":"No Loss statement confirmed for ' + htmlEscape_(submission.insuredName || '') + ' - ' + htmlEscape_(carrier) + ' policy ' + htmlEscape_(submission.policyNumber || '') + '","action":{"@type":"ViewAction","url":"https://www.BillLayneInsurance.com","name":"View Agency Website"}}</script>',
    '</head><body style="margin:0;padding:0;background-color:#f1f5f9;">',

    '<div style="display:none;white-space:nowrap;font:15px courier;color:#ffffff;line-height:0;width:600px!important;min-width:600px!important;max-width:600px!important;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>',
    '<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;color:#f1f5f9;line-height:1px;">your no loss statement is confirmed &#8212; copy attached for your records&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;</div>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#f1f5f9" style="background-color:#f1f5f9;"><tr><td align="center" style="padding:20px 10px;">',
    '<!--[if mso]><table align="center" border="0" cellspacing="0" cellpadding="0" width="600"><tr><td width="600"><![endif]-->',
    '<table cellpadding="0" cellspacing="0" border="0" width="600" class="email-container" style="max-width:600px;">',

    // CARD 1: HEADER
    '<tr><td style="padding-bottom:4px;"><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:16px 16px 0 0;border:1px solid #e2e8f0;">',
    '<tr><td style="height:4px;background-color:#003f87;font-size:0;line-height:0;border-radius:16px 16px 0 0;">&nbsp;</td></tr>',
    '<tr><td style="padding:20px 24px;" class="card-pad"><table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td align="center" valign="middle"><table cellpadding="0" cellspacing="0" border="0"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 12px;"><img src="' + logoUrl + '" width="160" alt="Bill Layne Insurance Agency" style="display:block;width:160px;max-width:160px;height:auto;"></td></tr></table></td></tr></table></td></tr>',
    '<tr><td style="padding:0 24px 14px;text-align:center;"><p style="margin:0;font-size:11px;color:#64748b;' + ff + 'letter-spacing:0.3px;">Statement of No Loss Confirmation &bull; ' + htmlEscape_(carrier) + ' &bull; Bill Layne Insurance Agency</p></td></tr>',
    '</table></td></tr>',

    // CARD 2: HERO GREEN
    '<tr><td style="padding-bottom:4px;"><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#059669" style="background-color:#059669;border-radius:16px;border:1px solid #e2e8f0;"><tr><td class="hero-pad" style="padding:36px 32px 40px;text-align:center;">',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px;"><tr><td bgcolor="#047857" style="background-color:#047857;border-radius:20px;padding:5px 16px;"><span style="font-size:11px;font-weight:700;color:#ffffff;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">&#9989; No Loss Statement Received</span></td></tr></table>',
    '<p style="margin:0 0 6px;font-size:13px;font-weight:600;color:rgba(255,255,255,0.80);' + ff + 'letter-spacing:0.5px;text-transform:uppercase;">Thank You</p>',
    '<p style="margin:0 0 8px;font-size:28px;font-weight:700;color:#ffffff;' + ff + 'line-height:1.2;">' + htmlEscape_(submission.insuredName || '') + '</p>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td bgcolor="#047857" style="background-color:#047857;border-radius:10px;padding:10px 20px;border:1px solid rgba(255,255,255,0.25);">',
    '<p style="margin:0;font-size:11px;font-weight:700;color:#C8A84E;' + ff + 'letter-spacing:1.2px;text-transform:uppercase;text-align:center;">Policy Number</p>',
    '<p style="margin:4px 0 0;font-size:14px;font-weight:600;color:#ffffff;' + ff + 'text-align:center;">' + htmlEscape_(submission.policyNumber || '') + '</p>',
    '<p style="margin:2px 0 0;font-size:12px;color:rgba(255,255,255,0.70);' + ff + 'text-align:center;">' + htmlEscape_(carrier) + ' &bull; ' + htmlEscape_(submission.policyType || 'Policy') + '</p>',
    '</td></tr></table>',
    '</td></tr></table></td></tr>',

    // CARD 3: BODY + DETAILS
    '<tr><td style="padding-bottom:4px;"><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;border:1px solid #e2e8f0;">',
    '<tr><td style="padding:28px 28px 0;" class="card-pad"><p style="margin:0 0 16px;font-size:15px;color:#334155;' + ff + 'line-height:1.65;">Your signed Statement of No Loss has been received and confirmed. A copy is attached to this email for your records &#8212; please save it in case you ever need it.</p></td></tr>',
    '<tr><td style="padding:0 28px 24px;" class="card-pad"><table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;"><tr><td style="padding:18px 20px;">',
    '<p style="margin:0 0 10px;font-size:10px;font-weight:700;color:#0369a1;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">Statement Details</p>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;' + ff + '">Confirmation #</td><td align="right" style="font-size:13px;font-weight:700;color:#0f172a;' + ff + '">' + htmlEscape_(submission.confirmationNumber) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;' + ff + '">Policy Number</td><td align="right" style="font-size:13px;color:#0f172a;' + ff + '">' + htmlEscape_(submission.policyNumber) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;' + ff + '">Insurance Company</td><td align="right" style="font-size:13px;color:#0f172a;' + ff + '">' + htmlEscape_(carrier) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:13px;color:#64748b;' + ff + '">Submitted</td><td align="right" style="font-size:13px;color:#0f172a;' + ff + '">' + htmlEscape_(localTime) + '</td></tr></table>',
    '</td></tr></table></td></tr></table></td></tr>',

    // CARD 4: WHAT HAPPENS NEXT
    '<tr><td style="padding-bottom:4px;"><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;border:1px solid #e2e8f0;border-left:4px solid #059669;"><tr><td style="padding:24px 28px;" class="card-pad">',
    '<p style="margin:0 0 4px;font-size:10px;font-weight:700;color:#64748b;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">Next Steps</p>',
    '<p style="margin:0 0 16px;font-size:20px;font-weight:700;color:#0f172a;' + ff + '">What Happens Next</p>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;"><tr><td style="padding:18px 20px;">',
    nlStep_(ff, '1', 'Securely stored', 'Your signed statement is saved in your policy file'),
    nlStep_(ff, '2', 'Forwarded to carrier', 'We\'ll send it to ' + htmlEscape_(carrier) + ' on your behalf'),
    nlStep_(ff, '3', 'Save your copy', 'The attached PDF is your receipt'),
    nlStepLast_(ff, '4', 'Nothing else needed', 'We\'ll reach out if anything changes'),
    '</td></tr></table>',
    '</td></tr></table></td></tr>',

    // CARD 5: CTA
    '<tr><td style="padding-bottom:4px;"><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;border:1px solid #e2e8f0;"><tr><td style="padding:24px 28px;" class="card-pad">',
    '<p style="margin:0 0 4px;font-size:15px;color:#334155;' + ff + 'line-height:1.6;">Thanks in advance,</p>',
    '<p style="margin:0 0 16px;font-size:15px;font-weight:700;color:#0f172a;' + ff + '">&mdash; Bill Layne</p>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td style="padding:0 4px;text-align:center;"><p style="margin:0;font-size:13px;color:#64748b;' + ff + 'line-height:1.6;font-style:italic;">Need a different format, additional documentation, or have questions about your policy? Reply here and I\'ll get back to you within the hour.</p></td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td class="btn-td" bgcolor="#003f87" style="background-color:#003f87;border-radius:8px;padding:14px 36px;"><a href="tel:3368351993" style="display:block;font-size:15px;font-weight:700;color:#ffffff;' + ff + 'text-decoration:none;text-align:center;">&#128222;&nbsp;&nbsp;Call (336) 835-1993</a></td></tr></table></td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top:12px;"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="border:2px solid #003f87;border-radius:8px;padding:12px 28px;"><a href="https://m.me/dollarbillagency?text=Hi%20Bill%2C%20I%20have%20a%20question%20about%20my%20no%20loss%20statement" target="_blank" style="display:block;font-size:14px;font-weight:700;color:#003f87;' + ff + 'text-decoration:none;text-align:center;">&#128172; Chat on Messenger</a></td></tr></table></td></tr><tr><td align="center" style="padding-top:6px;"><p style="margin:0;font-size:11px;color:#94a3b8;' + ff + '">Available Mon&ndash;Fri 9am&ndash;5pm &bull; We reply within 1 business hour</p></td></tr></table>',
    '</td></tr></table></td></tr>',

    // FOOTER
    '<tr><td><table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:0 0 16px 16px;border:1px solid #e2e8f0;border-top:none;"><tr><td style="padding:28px 24px;text-align:center;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" width="60" style="margin:0 auto 20px auto;"><tr><td style="height:3px;background:linear-gradient(90deg,#003f87,#C8A84E);font-size:0;line-height:0;">&nbsp;</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 12px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:6px;padding:6px 8px;"><img src="' + logoUrl + '" width="140" alt="Bill Layne Insurance Agency" style="display:block;width:140px;max-width:140px;height:auto;"></td></tr></table>',
    '<p style="margin:0 0 4px 0;font-size:14px;font-weight:700;color:#0f172a;' + ff + '">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 2px 0;font-size:12px;color:#64748b;' + ff + '">1283 N Bridge St &bull; Elkin, NC 28621</p>',
    '<p style="margin:0 0 2px 0;font-size:12px;color:#64748b;' + ff + '"><a href="tel:3368351993" style="color:#003f87;text-decoration:none;' + ff + '">(336)&nbsp;835&#8209;1993</a>&nbsp;&bull;&nbsp;<a href="mailto:Save@BillLayneInsurance.com" style="color:#003f87;text-decoration:none;' + ff + '">Save@BillLayneInsurance.com</a></p>',
    '<p style="margin:0 0 16px 0;font-size:12px;color:#64748b;' + ff + '"><a href="https://www.BillLayneInsurance.com" style="color:#003f87;text-decoration:none;' + ff + '">BillLayneInsurance.com</a>&nbsp;&bull;&nbsp;Est. 2005</p>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px auto;"><tr><td style="padding:0 6px;"><a href="https://facebook.com/dollarbillagency" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">Facebook</a></td><td style="color:#cbd5e1;font-size:11px;">|</td><td style="padding:0 6px;"><a href="https://youtube.com/@ncautoandhome" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">YouTube</a></td><td style="color:#cbd5e1;font-size:11px;">|</td><td style="padding:0 6px;"><a href="https://instagram.com/ncautoandhome" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">Instagram</a></td><td style="color:#cbd5e1;font-size:11px;">|</td><td style="padding:0 6px;"><a href="https://x.com/shopsavecompare" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">X</a></td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr><td style="background-color:#f8fafc;border-radius:8px;padding:8px 14px;border:1px solid #e2e8f0;"><table cellpadding="0" cellspacing="0" border="0"><tr><td valign="middle" style="padding-right:6px;"><img src="https://i.imgur.com/nDFmjxh.png" width="18" height="18" alt="Google" style="display:block;width:18px;height:18px;"></td><td valign="middle"><p style="margin:0;font-size:12px;font-weight:700;color:#0f172a;' + ff + '">4.9 &#11088;&#11088;&#11088;&#11088;&#11088; <span style="font-weight:400;color:#64748b;">100+ Google Reviews</span></p></td></tr></table></td></tr></table>',
    '<p style="margin:0 0 10px 0;font-size:11px;color:#64748b;' + ff + 'text-align:center;">Follow us on Facebook for tips, reminders &amp; updates &nbsp;&rarr;&nbsp;<a href="https://facebook.com/dollarbillagency" target="_blank" style="color:#003f87;font-weight:700;text-decoration:none;' + ff + '">facebook.com/dollarbillagency</a></p>',
    '<p style="margin:0 0 10px 0;font-size:11px;color:#64748b;' + ff + 'text-align:center;">&#128242; Want policy updates in Messenger? <a href="https://m.me/dollarbillagency?text=Yes%2C%20please%20send%20my%20policy%20updates%20via%20Messenger" target="_blank" style="color:#003f87;font-weight:700;text-decoration:none;' + ff + '">Tap here to connect &rarr;</a></p>',
    '<p style="margin:0;font-size:11px;color:#94a3b8;' + ff + '">To unsubscribe from agency communications, reply with UNSUBSCRIBE.</p>',
    '</td></tr></table></td></tr>',

    '</table>',
    '<!--[if mso]></td></tr></table><![endif]-->',
    '</td></tr></table>',
    '</body></html>'
  ].join('');

  MailApp.sendEmail({
    to: submission.email,
    subject: subject,
    name: runtime.fromName,
    replyTo: runtime.customerReplyTo,
    htmlBody: htmlBody,
    body: stripHtml_(htmlBody),
    attachments: [pdfFile.getBlob().setName(pdfFile.getName())]
  });

  return true;
}

function nlStep_(ff, num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:12px;"><tr><td width="36" valign="top"><table cellpadding="0" cellspacing="0" border="0" width="28" height="28"><tr><td width="28" height="28" align="center" valign="middle" bgcolor="#059669" style="background-color:#059669;border-radius:8px;font-size:13px;font-weight:700;color:#ffffff;' + ff + 'line-height:28px;">' + num + '</td></tr></table></td><td style="padding-left:8px;vertical-align:middle;"><p style="margin:0;font-size:14px;color:#334155;' + ff + 'line-height:1.5;"><strong style="color:#0f172a;">' + title + '</strong> &#8212; ' + desc + '</p></td></tr></table>';
}

function nlStepLast_(ff, num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td width="36" valign="top"><table cellpadding="0" cellspacing="0" border="0" width="28" height="28"><tr><td width="28" height="28" align="center" valign="middle" bgcolor="#059669" style="background-color:#059669;border-radius:8px;font-size:13px;font-weight:700;color:#ffffff;' + ff + 'line-height:28px;">' + num + '</td></tr></table></td><td style="padding-left:8px;vertical-align:middle;"><p style="margin:0;font-size:14px;color:#334155;' + ff + 'line-height:1.5;"><strong style="color:#0f172a;">' + title + '</strong> &#8212; ' + desc + '</p></td></tr></table>';
}

function sendCustomerConfirmationSms_(runtime, submission) {
  if (!hasTwilioConfig_(runtime) || !submission.phone) {
    return false;
  }

  const message = appendOptOut_(
    'Bill Layne Insurance received your Statement of No Loss for policy ' +
    submission.policyNumber + '. Confirmation #: ' + submission.confirmationNumber + '.'
  );

  sendSmsViaTwilio_(runtime, submission.phone, message);
  return true;
}

function sendSmsViaTwilio_(runtime, to, message) {
  if (!hasTwilioConfig_(runtime)) {
    throw new Error('Twilio settings are not configured.');
  }

  const payload = {
    To: normalizePhone_(to),
    Body: message
  };

  if (runtime.twilioMessagingServiceSid) {
    payload.MessagingServiceSid = runtime.twilioMessagingServiceSid;
  } else if (runtime.twilioFrom) {
    payload.From = runtime.twilioFrom;
  } else {
    throw new Error('Missing TWILIO_FROM or TWILIO_MESSAGING_SERVICE_SID.');
  }

  const response = UrlFetchApp.fetch(
    'https://api.twilio.com/2010-04-01/Accounts/' + runtime.twilioSid + '/Messages.json',
    {
      method: 'post',
      payload: payload,
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(runtime.twilioSid + ':' + runtime.twilioToken)
      },
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  const body = response.getContentText();
  let parsed = {};

  try {
    parsed = JSON.parse(body);
  } catch (error) {
    parsed = {};
  }

  if (code < 200 || code >= 300) {
    throw new Error('Twilio SMS failed: ' + (parsed.message || body || code));
  }

  return {
    sid: parsed.sid || '',
    status: parsed.status || ''
  };
}

function validateAgentPortalToken_(token) {
  if (!token) {
    throw new Error('Missing auth token.');
  }

  const props = PropertiesService.getScriptProperties();
  const secret = cleanText_(props.getProperty('AGENT_PORTAL_SECRET')) || APP_CONFIG.defaultPortalSecret;
  const today = new Date();
  const acceptedDates = [
    Utilities.formatDate(today, 'Etc/UTC', 'yyyy-MM-dd'),
    Utilities.formatDate(new Date(today.getTime() - 24 * 60 * 60 * 1000), 'Etc/UTC', 'yyyy-MM-dd'),
    Utilities.formatDate(new Date(today.getTime() + 24 * 60 * 60 * 1000), 'Etc/UTC', 'yyyy-MM-dd')
  ];

  const isValid = acceptedDates.some(function(dateValue) {
    return Utilities.base64Encode(secret + dateValue) === token;
  });

  if (!isValid) {
    throw new Error('Unauthorized agent portal request.');
  }
}

function buildArchiveObject_(submission) {
  return {
    confirmationNumber: submission.confirmationNumber,
    sessionId: submission.sessionId,
    insuredName: submission.insuredName,
    email: submission.email,
    phone: submission.phone,
    propertyAddress: submission.propertyAddress,
    city: submission.city,
    state: submission.state,
    zipCode: submission.zipCode,
    insuranceCompany: submission.insuranceCompany,
    policyNumber: submission.policyNumber,
    policyType: submission.policyType,
    amountPaid: submission.amountPaid,
    cancellationDate: submission.cancellationDate,
    reinstatementDate: submission.reinstatementDate,
    noLossConfirmation: submission.noLossConfirmation,
    dmvAcknowledgement: submission.dmvAcknowledgement,
    mortgageAcknowledgement: submission.mortgageAcknowledgement,
    agencyName: submission.agencyName,
    agencyAddress: submission.agencyAddress,
    agencyEmail: submission.agencyEmail,
    agencyPhone: submission.agencyPhone,
    agentName: submission.agentName,
    agentEmail: submission.agentEmail,
    signatureDateTime: submission.signatureDateTime,
    signatureTime: submission.signatureTime,
    submittedAt: submission.submittedAt,
    timezone: submission.timezone,
    ipAddress: submission.ipAddress,
    browserFingerprint: submission.browserFingerprint,
    deviceInfo: submission.deviceInfo,
    screenResolution: submission.screenResolution,
    userAgent: submission.userAgent,
    formVersion: submission.formVersion,
    submissionMethod: submission.submissionMethod,
    signatureMetadata: submission.signatureMetadata
  };
}


function hasTwilioConfig_(runtime) {
  return !!(runtime.twilioSid && runtime.twilioToken && (runtime.twilioFrom || runtime.twilioMessagingServiceSid));
}

function tryOptionalNotification_(fn, label) {
  try {
    return !!fn();
  } catch (error) {
    console.error('Optional notification failed (' + label + '): ' + error);
    return false;
  }
}

function appendOptOut_(message) {
  const text = cleanText_(message);
  if (!text) {
    return '';
  }
  return /reply\s+stop\s+to\s+opt\s+out\.?$/i.test(text) ? text : text + '\n\nReply STOP to opt out.';
}

function normalizePhone_(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  if (!digits) {
    return '';
  }
  if (digits.length === 10) {
    return '+1' + digits;
  }
  if (digits.length === 11 && digits.charAt(0) === '1') {
    return '+' + digits;
  }
  return digits.charAt(0) === '+' ? digits : '+' + digits;
}

function requireField_(value, fieldName) {
  const cleaned = cleanText_(value);
  if (!cleaned) {
    throw new Error('Missing required field: ' + fieldName);
  }
  return cleaned;
}

function truthyToYesNo_(value, forceYesForTruthy) {
  if (value === true || value === 'true' || value === 'on' || value === 'Yes' || value === 'yes') {
    return 'Yes';
  }
  if (forceYesForTruthy && cleanText_(value)) {
    return 'Yes';
  }
  return '';
}

function parseMaybeJson_(value) {
  const cleaned = cleanText_(value);
  if (!cleaned) {
    return {};
  }

  try {
    return JSON.parse(cleaned);
  } catch (error) {
    return { raw: cleaned };
  }
}

function safeDate_(value) {
  const date = new Date(value);
  return isNaN(date.getTime()) ? new Date() : date;
}

function formatDateTime_(value) {
  return Utilities.formatDate(safeDate_(value), APP_CONFIG.defaultTimeZone, 'MMMM d, yyyy h:mm a');
}

function findOrCreateFolder_(parentFolder, childName) {
  const matches = parentFolder.getFoldersByName(childName);
  return matches.hasNext() ? matches.next() : parentFolder.createFolder(childName);
}

function sanitizeFileName_(name) {
  return String(name || 'submission')
    .replace(/[\\/:*?"<>|#%&{}$!'@+=`]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function generateFallbackConfirmation_() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, APP_CONFIG.defaultTimeZone, 'yyyyMMdd');
  const randomPart = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return 'NOL-' + datePart + '-' + randomPart;
}

function cleanText_(value) {
  return String(value || '').trim();
}

function isValidEmail_(value) {
  const email = cleanText_(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function stripHtml_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .trim();
}

function htmlEscape_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
