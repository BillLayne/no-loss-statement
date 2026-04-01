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

  // Set narrow margins for single-page fit
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(54);
  body.setMarginRight(54);

  // Agency header - compact
  var header = body.appendParagraph(submission.agencyName || APP_CONFIG.agencyName);
  header.setFontSize(14).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var subHeader = body.appendParagraph('STATEMENT OF NO LOSS');
  subHeader.setFontSize(11).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2);

  var confLine = body.appendParagraph('Confirmation: ' + submission.confirmationNumber + '  |  ' + formatDateTime_(submission.submittedAt));
  confLine.setFontSize(8).setForegroundColor('#666666').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(4);

  body.appendHorizontalRule();

  // Policy Info - two-column style using compact key-value
  appendCompactHeading_(body, 'POLICY INFORMATION');
  appendCompactKV_(body, 'Insurance Company', submission.insuranceCompany);
  appendCompactKV_(body, 'Policy Number', submission.policyNumber);
  appendCompactKV_(body, 'Policy Type', submission.policyType);
  appendCompactKV_(body, 'Amount to Reinstate', submission.amountPaid || 'N/A');
  appendCompactKV_(body, 'Cancellation / Lapse Date', submission.cancellationDate);
  appendCompactKV_(body, 'Requested Reinstatement', submission.reinstatementDate);

  // Insured Info
  appendCompactHeading_(body, 'INSURED INFORMATION');
  appendCompactKV_(body, 'Name', submission.insuredName);
  appendCompactKV_(body, 'Phone', submission.phone);
  appendCompactKV_(body, 'Email', submission.email || 'N/A');
  appendCompactKV_(body, 'Address', submission.propertyAddress + ', ' + submission.city + ', ' + submission.state + ' ' + submission.zipCode);

  // Acknowledgements - inline
  appendCompactHeading_(body, 'ACKNOWLEDGEMENTS');
  appendCompactKV_(body, 'No Loss Confirmed', submission.noLossConfirmation);
  appendCompactKV_(body, 'DMV Acknowledgement', submission.dmvAcknowledgement || 'No');
  appendCompactKV_(body, 'Mortgage Acknowledgement', submission.mortgageAcknowledgement || 'No');

  // Statement text - single condensed paragraph
  appendCompactHeading_(body, 'STATEMENT');
  var stmtText = 'I, ' + submission.insuredName + ', state that neither I nor any other person covered by this policy has had a claim or loss or been involved in an accident since the cancellation or expiration of the policy. I understand the insurance company is relying on this statement to reinstate my policy with no lapse in coverage. If a claim occurred during the no loss period, the reinstatement is null and void. If my payment is not honored, the reinstatement is null and void and no coverage shall exist.';
  var stmt = body.appendParagraph(stmtText);
  stmt.setFontSize(9).setSpacingAfter(6).setLineSpacing(1.1);

  // Signature block - compact
  body.appendHorizontalRule();
  var sigLabel = body.appendParagraph('ELECTRONIC SIGNATURE');
  sigLabel.setFontSize(8).setBold(true).setForegroundColor('#333333').setSpacingAfter(2);

  var sigInfo = body.appendParagraph('Signed by: ' + submission.insuredName + '  |  ' + (submission.signatureDateTime || formatDateTime_(submission.submittedAt)) + '  |  IP: ' + (submission.ipAddress || 'N/A'));
  sigInfo.setFontSize(7).setForegroundColor('#666666').setSpacingAfter(4);

  body.appendImage(signatureBlob).setWidth(200);

  doc.saveAndClose();

  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs(MimeType.PDF).setName(docName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);

  return pdfFile;
}

function appendCompactHeading_(body, title) {
  var p = body.appendParagraph(title);
  p.setFontSize(9).setBold(true).setForegroundColor('#003f87').setSpacingBefore(6).setSpacingAfter(2);
}

function appendCompactKV_(body, label, value) {
  var p = body.appendParagraph(label + ':  ' + (value || ''));
  p.setFontSize(9).setSpacingAfter(1).setSpacingBefore(0);
  // Bold just the label portion
  var text = p.editAsText();
  text.setBold(0, label.length, true);
  text.setBold(label.length, p.getText().length - 1, false);
  text.setForegroundColor(0, label.length, '#333333');
  text.setForegroundColor(label.length + 1, p.getText().length - 1, '#000000');
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

  var subject = '\u2705 Statement of No Loss Received \u2014 ' + htmlEscape_(submission.insuredName) + ' \u2014 ' + htmlEscape_(carrier);
  var logoUrl = 'https://i.imgur.com/lxu9nfT.png';
  var carrierSlug = carrier.toLowerCase().replace(/[^a-z0-9]+/g, '_');
  var clientSlug = (submission.insuredName || '').toLowerCase().replace(/[^a-z0-9]+/g, '_');

  var htmlBody = [
    '<!DOCTYPE html><html lang="en" xmlns="http://www.w3.org/1999/xhtml"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="x-apple-disable-message-reformatting"><title>Statement of No Loss Received</title>',
    '<style>body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0;mso-table-rspace:0}img{-ms-interpolation-mode:bicubic;border:0;height:auto;line-height:100%;outline:none;text-decoration:none}body{margin:0;padding:0;width:100%!important;background-color:#f1f5f9}@media only screen and (max-width:620px){.email-container{width:100%!important;padding:0 12px!important}.hero-pad{padding:28px 20px 32px!important}.card-pad{padding:20px 16px!important}.btn-td{padding:14px 24px!important}}</style>',
    '</head><body style="margin:0;padding:0;background-color:#f1f5f9;font-family:Arial,\'Helvetica Neue\',Helvetica,sans-serif;">',

    '<div style="display:none;font-size:1px;color:#f1f5f9;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;">your no loss statement is confirmed &#8212; copy attached for your records</div>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#f1f5f9" style="background-color:#f1f5f9;"><tr><td align="center" style="padding:24px 16px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="600" class="email-container" style="width:600px;max-width:600px;margin:0 auto;">',

    // ── CARD 1: HEADER ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fafafa;border-radius:16px 16px 0 0;border:1px solid #e2e8f0;border-bottom:none;">',
    '<tr><td style="height:4px;background-color:#003f87;font-size:0;line-height:0;border-radius:16px 16px 0 0;">&nbsp;</td></tr>',
    '<tr><td style="padding:20px 24px;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>',
    '<td align="center" valign="middle">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 12px;">',
    '<img src="' + logoUrl + '" width="160" alt="Bill Layne Insurance Agency" style="display:block;width:160px;max-width:160px;height:auto;">',
    '</td></tr></table>',
    '</td>',
    '</tr></table>',
    '</td></tr>',
    '<tr><td style="padding:0 24px 14px;text-align:center;">',
    '<p style="margin:0;font-size:11px;color:#64748b;font-family:Arial,sans-serif;letter-spacing:0.3px;">Statement of No Loss Confirmation &bull; ' + htmlEscape_(carrier) + ' &bull; Bill Layne Insurance Agency</p>',
    '</td></tr>',
    '</table></td></tr>',

    // ── CARD 2: HERO GREEN ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td class="hero-pad" style="padding:36px 32px 40px;background-color:#059669;text-align:center;">',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px;"><tr><td style="background-color:#047857;border-radius:20px;padding:5px 16px;">',
    '<span style="font-size:11px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">&#9989; No Loss Statement Received</span>',
    '</td></tr></table>',

    '<p style="margin:0 0 6px;font-size:13px;font-weight:600;color:rgba(255,255,255,0.80);font-family:Arial,sans-serif;letter-spacing:0.5px;text-transform:uppercase;">Thank You</p>',
    '<p style="margin:0 0 8px;font-size:28px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;line-height:1.2;">' + htmlEscape_(submission.insuredName) + '</p>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>',
    '<td style="background-color:#047857;border-radius:10px;padding:10px 20px;border:1px solid rgba(255,255,255,0.25);">',
    '<p style="margin:0;font-size:11px;font-weight:700;color:#C8A84E;font-family:Arial,sans-serif;letter-spacing:1.2px;text-transform:uppercase;text-align:center;">Policy Number</p>',
    '<p style="margin:4px 0 0;font-size:14px;font-weight:600;color:#ffffff;font-family:Arial,sans-serif;text-align:center;">' + htmlEscape_(submission.policyNumber) + '</p>',
    '<p style="margin:2px 0 0;font-size:12px;color:rgba(255,255,255,0.70);font-family:Arial,sans-serif;text-align:center;">' + htmlEscape_(carrier) + ' &bull; ' + htmlEscape_(submission.policyType || 'Policy') + '</p>',
    '</td></tr></table>',

    '</td></tr></table></td></tr>',

    // ── CARD 3: BODY + DETAILS ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:28px 28px 0;" class="card-pad">',
    '<p style="margin:0 0 16px;font-size:15px;color:#334155;font-family:Arial,sans-serif;line-height:1.65;">Your signed Statement of No Loss has been received and confirmed. A copy is attached to this email for your records &#8212; please save it in case you ever need it.</p>',
    '</td></tr>',

    '<tr><td style="padding:0 28px 20px;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;">',
    '<tr><td style="padding:18px 20px;">',
    '<p style="margin:0 0 10px;font-size:10px;font-weight:700;color:#0369a1;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">Statement Details</p>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">Confirmation #</td><td align="right" style="font-size:13px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">' + htmlEscape_(submission.confirmationNumber) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">Policy Number</td><td align="right" style="font-size:13px;color:#0f172a;font-family:Arial,sans-serif;">' + htmlEscape_(submission.policyNumber) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">Insurance Company</td><td align="right" style="font-size:13px;color:#0f172a;font-family:Arial,sans-serif;">' + htmlEscape_(carrier) + '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">Submitted</td><td align="right" style="font-size:13px;color:#0f172a;font-family:Arial,sans-serif;">' + htmlEscape_(localTime) + '</td></tr></table>',

    '</td></tr></table>',
    '</td></tr></table></td></tr>',

    // ── CARD 4: WHAT HAPPENS NEXT ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:24px 28px;" class="card-pad">',
    '<p style="margin:0 0 4px;font-size:10px;font-weight:700;color:#64748b;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">Next Steps</p>',
    '<p style="margin:0 0 16px;font-size:20px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">What Happens Next</p>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;">',
    '<tr><td style="padding:18px 20px;">',

    buildStep_(1, 'Securely stored', 'Your signed statement is saved in your policy file'),
    buildStep_(2, 'Forwarded to carrier', 'We\'ll send it to ' + htmlEscape_(carrier) + ' on your behalf'),
    buildStep_(3, 'Save your copy', 'The attached PDF is your receipt'),
    buildStepLast_(4, 'Nothing else needed', 'We\'ll reach out if anything changes'),

    '</td></tr></table>',
    '</td></tr></table></td></tr>',

    // ── CARD 5: CTA ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:24px 28px;" class="card-pad">',

    '<p style="margin:0 0 4px;font-size:15px;color:#334155;font-family:Arial,sans-serif;line-height:1.6;">Thanks in advance,</p>',
    '<p style="margin:0 0 16px;font-size:15px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">&mdash; Bill Layne</p>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td style="padding:0 4px;text-align:center;">',
    '<p style="margin:0;font-size:13px;color:#64748b;font-family:Arial,sans-serif;line-height:1.6;font-style:italic;">Need a different format, additional documentation, or have questions about your policy? Reply here and I\'ll get back to you within the hour.</p>',
    '</td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td align="center">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td class="btn-td" style="background-color:#003f87;border-radius:12px;padding:14px 36px;">',
    '<a href="tel:3368351993" style="display:block;font-size:15px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;text-decoration:none;text-align:center;">Call (336) 835-1993</a>',
    '</td></tr></table></td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top:12px;"><tr><td align="center">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="border:2px solid #003f87;border-radius:12px;padding:12px 28px;">',
    '<a href="https://m.me/dollarbillagency?text=Hi%20Bill%2C%20I%20have%20a%20question%20about%20my%20no%20loss%20statement" target="_blank" style="display:block;font-size:14px;font-weight:700;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;text-align:center;">&#128172; Chat on Messenger</a>',
    '</td></tr></table></td></tr>',
    '<tr><td align="center" style="padding-top:6px;">',
    '<p style="margin:0;font-size:11px;color:#94a3b8;font-family:Arial,sans-serif;">Available Mon&ndash;Fri 9am&ndash;5pm &bull; We reply within 1 business hour</p>',
    '</td></tr></table>',

    '</td></tr></table></td></tr>',

    // ── FOOTER ──
    '<tr><td>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:0 0 16px 16px;border:1px solid #e2e8f0;border-top:none;">',
    '<tr><td style="padding:28px 24px;text-align:center;" class="card-pad">',

    '<table cellpadding="0" cellspacing="0" border="0" width="60" style="margin:0 auto 20px auto;"><tr><td style="height:3px;background-color:#003f87;font-size:0;line-height:0;">&nbsp;</td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 12px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 12px;">',
    '<img src="' + logoUrl + '" width="140" alt="Bill Layne Insurance Agency" style="display:block;width:140px;max-width:140px;height:auto;">',
    '</td></tr></table>',

    '<p style="margin:0 0 4px;font-size:14px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 2px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;">1283 N Bridge St &bull; Elkin, NC 28621</p>',
    '<p style="margin:0 0 2px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;"><a href="tel:3368351993" style="color:#64748b;text-decoration:none;">(336) 835-1993</a> &bull; <a href="mailto:Save@BillLayneInsurance.com" style="color:#64748b;text-decoration:none;">Save@BillLayneInsurance.com</a></p>',
    '<p style="margin:0 0 14px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;"><a href="https://www.BillLayneInsurance.com" style="color:#64748b;text-decoration:none;">www.BillLayneInsurance.com</a> &bull; Est. 2005</p>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.facebook.com/dollarbillagency" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">Facebook</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.youtube.com/@ncautoandhome" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">YouTube</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.instagram.com/ncautoandhome" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">Instagram</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://twitter.com/shopsavecompare" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">X</a></td></tr></table></td>',
    '</tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr><td style="background-color:#f8fafc;border-radius:8px;padding:8px 14px;border:1px solid #e2e8f0;">',
    '<p style="margin:0;font-size:12px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">4.9 &#11088;&#11088;&#11088;&#11088;&#11088; <span style="font-weight:400;color:#64748b;">100+ Google Reviews</span></p>',
    '</td></tr></table>',

    '<p style="margin:0 0 10px;font-size:11px;color:#64748b;font-family:Arial,sans-serif;text-align:center;">Follow us on Facebook for tips, reminders &amp; updates &nbsp;&rarr;&nbsp;<a href="https://facebook.com/dollarbillagency" target="_blank" style="color:#003f87;font-weight:700;text-decoration:none;">facebook.com/dollarbillagency</a></p>',

    '<p style="margin:0;font-size:11px;color:#94a3b8;font-family:Arial,sans-serif;text-align:center;">You\'re receiving this because you\'re a valued client of Bill Layne Insurance Agency.<br>&copy; 2026 Bill Layne Insurance Agency. All rights reserved.</p>',

    '</td></tr></table></td></tr>',

    '</table>',
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

function buildStep_(num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:12px;"><tr><td width="36" valign="top"><table cellpadding="0" cellspacing="0" border="0" width="28" height="28"><tr><td width="28" height="28" align="center" valign="middle" style="background-color:#059669;border-radius:8px;font-size:13px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;line-height:28px;">' + num + '</td></tr></table></td><td style="padding-left:8px;vertical-align:middle;"><p style="margin:0;font-size:14px;color:#334155;font-family:Arial,sans-serif;line-height:1.5;"><strong style="color:#0f172a;">' + title + '</strong> &#8212; ' + desc + '</p></td></tr></table>';
}

function buildStepLast_(num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td width="36" valign="top"><table cellpadding="0" cellspacing="0" border="0" width="28" height="28"><tr><td width="28" height="28" align="center" valign="middle" style="background-color:#059669;border-radius:8px;font-size:13px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;line-height:28px;">' + num + '</td></tr></table></td><td style="padding-left:8px;vertical-align:middle;"><p style="margin:0;font-size:14px;color:#334155;font-family:Arial,sans-serif;line-height:1.5;"><strong style="color:#0f172a;">' + title + '</strong> &#8212; ' + desc + '</p></td></tr></table>';
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
