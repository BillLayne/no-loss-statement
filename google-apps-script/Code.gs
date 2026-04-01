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
  const docName = submission.confirmationNumber + ' - Statement of No Loss';
  const doc = DocumentApp.create(docName);
  const docId = doc.getId();
  const body = doc.getBody();

  body.appendParagraph(submission.agencyName || APP_CONFIG.agencyName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Statement of No Loss')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Confirmation #: ' + submission.confirmationNumber);
  body.appendParagraph('Submitted: ' + formatDateTime_(submission.submittedAt));
  body.appendHorizontalRule();

  appendSectionHeading_(body, 'Policy Information');
  appendKeyValue_(body, 'Insurance Company', submission.insuranceCompany);
  appendKeyValue_(body, 'Policy Number', submission.policyNumber);
  appendKeyValue_(body, 'Policy Type', submission.policyType);
  appendKeyValue_(body, 'Amount to Reinstate', submission.amountPaid || 'Not provided');
  appendKeyValue_(body, 'Cancellation / Lapse Date', submission.cancellationDate);
  appendKeyValue_(body, 'Requested Reinstatement', submission.reinstatementDate);

  appendSectionHeading_(body, 'Insured Information');
  appendKeyValue_(body, 'Insured Name', submission.insuredName);
  appendKeyValue_(body, 'Email', submission.email || 'Not provided');
  appendKeyValue_(body, 'Phone', submission.phone);
  appendKeyValue_(body, 'Property / Garaging Address', submission.propertyAddress);
  appendKeyValue_(body, 'City / State / ZIP', submission.city + ', ' + submission.state + ' ' + submission.zipCode);

  appendSectionHeading_(body, 'Acknowledgements');
  appendKeyValue_(body, 'No Loss Confirmation', submission.noLossConfirmation);
  appendKeyValue_(body, 'DMV Acknowledgement', submission.dmvAcknowledgement || 'No');
  appendKeyValue_(body, 'Mortgage Acknowledgement', submission.mortgageAcknowledgement || 'No');

  appendSectionHeading_(body, 'Statement Text');
  body.appendParagraph(
    'I, ' + submission.insuredName + ', state that neither I nor any other person covered by this policy has had a claim or loss or been involved in an accident since the cancellation or expiration of the policy wherein this policy may apply.'
  );
  body.appendParagraph(
    'I understand that the insurance company is relying on this Statement of No Loss as an inducement to reinstate my policy with no lapse in coverage. I understand that if a claim, loss, or accident occurred during the no loss period, the reinstatement is null and void and coverage may be denied.'
  );
  body.appendParagraph(
    'I agree that if my payment for this reinstatement is not honored for any reason, the reinstatement is null and void and no coverage shall exist under this policy.'
  );

  appendSectionHeading_(body, 'Electronic Signature');
  body.appendParagraph('Signed by: ' + submission.insuredName);
  body.appendParagraph('Signature captured: ' + (submission.signatureDateTime || formatDateTime_(submission.submittedAt)));
  body.appendParagraph('IP Address: ' + (submission.ipAddress || 'Not captured'));
  body.appendParagraph('Device Info: ' + (submission.deviceInfo || 'Not captured'));
  body.appendParagraph('Browser Fingerprint: ' + (submission.browserFingerprint || 'Not captured'));
  body.appendParagraph('Submission Method: ' + (submission.submissionMethod || 'Online Portal'));
  body.appendParagraph('Form Version: ' + (submission.formVersion || 'Unknown'));
  body.appendParagraph('');
  body.appendImage(signatureBlob).setWidth(260);

  appendSectionHeading_(body, 'Metadata');
  body.appendParagraph(JSON.stringify(buildArchiveObject_(submission), null, 2))
    .setFontFamily('Courier New')
    .setFontSize(8);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(docId);
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(docName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);

  return pdfFile;
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
  var localTime = '';
  try {
    localTime = Utilities.formatDate(new Date(submission.submittedAt), APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch(e) {
    localTime = submission.submittedAt || new Date().toLocaleString();
  }

  var subject = '\u2705 Statement of No Loss Received \u2014 ' + submission.confirmationNumber + ' | Bill Layne Insurance';
  var logoUrl = 'https://i.imgur.com/lxu9nfT.png';

  var htmlBody = [
    '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background-color:#f1f5f9;-webkit-text-size-adjust:100%;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:600px;margin:0 auto;">',

    '<tr><td style="padding:0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#003f87;background:linear-gradient(135deg,#003f87 0%,#0076d3 100%);border-radius:0 0 16px 16px;">',
    '<tr><td style="padding:36px 30px 28px;text-align:center;">',
    '<img src="' + logoUrl + '" alt="Bill Layne Insurance" width="180" height="45" style="display:block;margin:0 auto 16px;max-width:180px;height:auto;border:0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:#1a5296;border-radius:20px;padding:6px 16px;"><span style="font-family:Arial,sans-serif;font-size:13px;color:#ffffff;">&#10003; No Loss Statement Received</span></td></tr></table>',
    '</td></tr></table></td></tr>',

    '<tr><td style="padding:20px 16px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ffffff;border-radius:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">',

    '<tr><td style="padding:28px 28px 0;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:22px;font-weight:700;color:#0f2744;">Thank you, ' + htmlEscape_(firstName) + '!</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:15px;color:#64748b;line-height:1.6;">We have received your signed Statement of No Loss and appreciate you completing this promptly. A copy of your signed statement is attached to this email for your records.</p>',
    '</td></tr>',

    '<tr><td style="padding:20px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;">',
    '<tr><td style="padding:16px 20px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">',
    '<tr><td style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:0.5px;padding-bottom:8px;">Statement Details</td></tr>',
    '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Confirmation #</td><td style="font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:#0f2744;text-align:right;">' + htmlEscape_(submission.confirmationNumber) + '</td></tr></table></td></tr>',
    '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Policy Number</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;">' + htmlEscape_(submission.policyNumber) + '</td></tr></table></td></tr>',
    '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Insurance Company</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;">' + htmlEscape_(submission.insuranceCompany) + '</td></tr></table></td></tr>',
    '<tr><td><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Submitted</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;">' + htmlEscape_(localTime) + '</td></tr></table></td></tr>',
    '</table></td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;">',
    '<tr><td style="padding:16px 20px;">',
    '<p style="margin:0 0 8px;font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;">What Happens Next</p>',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0">',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128203; Your signed statement is securely stored</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128222; We will forward it to your insurance company</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128274; Keep the attached PDF copy for your records</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128222; We will contact you if anything else is needed</td></tr>',
    '</table></td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fff7ed;border-radius:12px;border:1px solid #fed7aa;">',
    '<tr><td style="padding:16px 20px;text-align:center;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:#9a3412;">Have questions?</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:14px;color:#78350f;">Call us at <a href="tel:3368351993" style="color:#0076d3;text-decoration:none;font-weight:700;">(336) 835-1993</a></p>',
    '</td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="text-align:center;">',
    '<a href="https://www.billlayneinsurance.com" target="_blank" style="display:inline-block;background-color:#0076d3;color:#ffffff;font-family:Arial,sans-serif;font-size:15px;font-weight:700;text-decoration:none;padding:14px 36px;border-radius:12px;">Visit Our Website</a>',
    '</td></tr></table></td></tr>',

    '</table></td></tr>',

    '<tr><td style="padding:20px 16px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#0f172a;border-radius:16px;">',
    '<tr><td style="padding:28px 28px 20px;text-align:center;">',
    '<img src="' + logoUrl + '" alt="Bill Layne Insurance" width="140" height="35" style="display:block;margin:0 auto 12px;max-width:140px;height:auto;border:0;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:14px;color:#e2e8f0;">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;">1283 N Bridge St, Elkin, NC 28621</p>',
    '<p style="margin:0 0 12px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;"><a href="tel:3368351993" style="color:#60a5fa;text-decoration:none;">(336) 835-1993</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="mailto:docs@billlayneinsurance.com" style="color:#60a5fa;text-decoration:none;">docs@billlayneinsurance.com</a></p>',
    '</td></tr>',
    '<tr><td style="padding:0 28px 20px;text-align:center;"><p style="margin:0;font-family:Arial,sans-serif;font-size:11px;color:#475569;">&copy; 2026 Bill Layne Insurance Agency. All rights reserved.</p></td></tr>',
    '</table></td></tr>',

    '<tr><td style="padding:20px 0;">&nbsp;</td></tr>',
    '</table></body></html>'
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

function appendSectionHeading_(body, title) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING3);
}

function appendKeyValue_(body, label, value) {
  body.appendParagraph(label + ': ' + (value || ''));
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
