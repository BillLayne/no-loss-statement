# No Loss Statement Google Apps Script Backend

This folder contains a replacement backend for `mynolossform.com` and its agent prefill portal.

It supports both flows:

1. Customer submits the signed Statement of No Loss form
2. Agent portal sends the customer a secure SMS link

## What The Script Does

For customer submissions, `Code.gs` will:

- accept the current form payload from `index.html`
- save a signature image to Google Drive
- save a JSON archive of the submission
- generate a PDF copy of the signed statement
- email the office a copy
- email the customer a confirmation if an email address was provided
- email the agent a copy if `agentEmail` was included in the prefilled link
- send the customer a confirmation SMS if Twilio settings are configured

For agent portal SMS sends, it will:

- validate the agent portal auth token
- call Twilio server-side
- append `Reply STOP to opt out.` if it is missing

## Required Script Properties

In Apps Script `Project Settings`, add:

`UPLOADS_ROOT_FOLDER_ID`
Google Drive folder ID where statement folders should be created.

`OFFICE_EMAILS`
Comma-separated office recipients.
Example:
`billlayneinsurance@gmail.com,save@billlayneinsurance.com`

## Optional Script Properties

`FROM_NAME`
Example: `Bill Layne Insurance Agency`

`CUSTOMER_REPLY_TO`
Reply-to address used on outgoing emails.

`TIMEZONE`
Example: `America/New_York`

`AGENT_PORTAL_SECRET`
Defaults to the current portal secret in `agent-portal.html`.
Set this if you want to rotate the secret later.

## Twilio Properties

Add these if you want SMS features to work:

`TWILIO_SID`

`TWILIO_TOKEN`

One of:

`TWILIO_FROM`

or

`TWILIO_MESSAGING_SERVICE_SID`

If Twilio properties are missing, the form submission still works, but SMS confirmations and agent-portal SMS sends will not.

## Deploy

1. Go to [Google Apps Script](https://script.google.com/).
2. Create a new project.
3. Replace the default file with the contents of [Code.gs](C:/Users/bill/OneDrive/Documents/Playground/no-loss-statement/google-apps-script/Code.gs).
4. Add the script properties above.
5. Click `Deploy` -> `New deployment`.
6. Choose `Web app`.
7. Set `Execute as` to `Me`.
8. Set access to `Anyone`.
9. Deploy and authorize it.
10. Copy the new `/exec` URL.

## Update The Site

After you deploy the new Apps Script:

- replace `SCRIPT_URL` in [index.html](C:/Users/bill/OneDrive/Documents/Playground/no-loss-statement/index.html)
- replace `GOOGLE_SCRIPT_URL` in [agent-portal.html](C:/Users/bill/OneDrive/Documents/Playground/no-loss-statement/agent-portal.html)

## Recommended Test

1. Open the public form and submit a test statement.
2. Confirm the Drive folder, signature image, JSON archive, and PDF were created.
3. Confirm the office email arrives with the PDF.
4. Confirm the customer email arrives if you entered a customer email.
5. Confirm the agent copy arrives if you included `agentEmail` in the prefill link.
6. Test the agent portal SMS button if Twilio is configured.

## Why It Broke

The old Apps Script web app URL currently hardcoded in the repo returns `404 Not Found`, which means the previous deployment was removed or no longer exists at that URL.
