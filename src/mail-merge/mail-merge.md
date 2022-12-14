# Mail merge with attachment

## Introduction ##

We would like to send multiple emails to all of participants who purchased a ticket for our upcoming DevFest.

We tried multiple options suggested on the internet but either they are paid or could not solve the problem statement we had.

First part was to get multiple QR code created by a cell value from the Google Sheet.

This was pretty straight forward and was solved by using a `=IMAGE(image-url)` formula in a cell.

To get qr code we used the online utility as api from https://quickchart.io/documentation/qr-codes/

That would simply create QR per record but how do we send these many emails?

Using mail merge as described here https://workspace.google.com/marketplace/app/mail_merge/218858140171 we were able to get a basic app script to send out builk emails.

> You can download the attached .csv file for your test and use the attached .gs script as the related sheet's app script.

> IMPORTANT: The csv has a QR column without value in it. 
> Add the following formula to generate your own QR 
> =IMAGE("https://quickchart.io/qr?text=" & ENCODEURL(cell-number))

## Send an email using mail merge ##

You can follow step by step guide on how to setup your draft email template (using your Gmail account) https://workspace.google.com/marketplace/app/mail_merge/218858140171

## Problem ##

We were able to send emails but noticed the image is not getting attached nor getting set as inline image in body of the email.

## Solution ##

After fiddling the internet we notice that since app scripts are just javascript run time, how about if we can generate the image on fly and try to attach to the outgoing email.

Thanks to this stack overflow answer here: https://stackoverflow.com/questions/49199116/add-image-to-google-sheets-email-app-script

That gave us the hint on how we can also include a blob as attachment before sending out the email.

```
// IMPORTANT: This will generate a QR Code image and get blob to add as attachment
    var qrBlob = UrlFetchApp.fetch("https://quickchart.io/qr?text=" + row[QR_TEXT_COL]).getBlob();

    // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
    // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
    // Uncomment advanced parameters as needed (see docs for limitations)
    GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
        htmlBody: msgObj.html,
        // bcc: 'a.bbc@email.com',
        // cc: 'a.cc@email.com',
        // from: 'an.alias@email.com',
        // name: 'name of the sender',
        // replyTo: 'a.reply@email.com',
        // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
        attachments: [qrBlob],
        inlineImages: emailTemplate.inlineImages
});
```

> IMPORTNANT: This solution now DOES NOT make use of the QR column in the test csv

> IMPORTANT: In the code snippet above the `QR_TEXT_COL` is a const defined in the script itself. Review code.gs in this folder for more details.

## Tips about the script ##

- Your test worksheet should include `Email` as column header with values
- Your test worksheet should include `Email Sent` as column header (value can be left empty)
- Your test worksheet should include `Transaction Id` as column header with values.
- You can replace these column header names to the name of your choice but make sure you also replace the names in the code.gs in your sheet's app script.