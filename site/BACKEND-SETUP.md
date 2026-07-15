# Backend Setup — Contact Form + Review Moderation

Your site now relies on one shared, free backend (Google Apps Script) for two things:

1. **Contact form** → logs every submission to a Google Sheet, and Netlify emails you a notification.
2. **Client reviews** → a submitted review is emailed to you with an **Approve** button. Nothing appears on the website until you click it. Approved reviews are then pulled from the Sheet and shown automatically.

This takes about 15 minutes, entirely inside your free Google account.

---

## Part 1 — Email notifications for the contact form (Netlify Forms)

1. Deploy your site to Netlify.
2. Netlify dashboard → **Site settings → Forms → Form notifications → Add notification → Email notification**.
3. Set the email to `infonuvagold@gmail.com`. Save.

Every contact form submission now emails that inbox automatically — no code involved.

---

## Part 2 — The shared backend (Google Sheet + Apps Script)

### Step 1: Create the Google Sheet
1. Go to [sheets.google.com](https://sheets.google.com) → create a new blank sheet.
2. Name it `NUVA-Gold Website Data`. You don't need to add any tabs or headers manually — the script creates and labels them the first time it runs.

### Step 2: Add the script
1. In the sheet, go to **Extensions → Apps Script**.
2. Delete any starter code, and paste this in full:

```javascript
var ADMIN_EMAIL = "infonuvagold@gmail.com";

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.formType === 'review') {
    return handleReviewSubmission(ss, data);
  }
  return handleContactSubmission(ss, data);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action;

  if (action === 'approve' && e.parameter.id) {
    return handleApprove(ss, e.parameter.id);
  }
  if (action === 'delete' && e.parameter.id) {
    return handleDelete(ss, e.parameter.id);
  }
  if (action === 'manage') {
    return handleManage(ss);
  }
  return getApprovedReviews(ss);
}

function handleContactSubmission(ss, data) {
  var sheet = ss.getSheetByName('Contact');
  if (!sheet) {
    sheet = ss.insertSheet('Contact');
    sheet.appendRow(['Timestamp', 'Name', 'Company', 'Email', 'Service', 'Message']);
  }
  sheet.appendRow([new Date(), data.name || '', data.company || '', data.email || '', data.service || '', data.message || '']);
  return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

function handleReviewSubmission(ss, data) {
  var sheet = ss.getSheetByName('Reviews');
  if (!sheet) {
    sheet = ss.insertSheet('Reviews');
    sheet.appendRow(['ID', 'Timestamp', 'Name', 'Rating', 'Text', 'Approved']);
  }

  var id = Utilities.getUuid();
  var rating = Number(data.rating) || 0;
  sheet.appendRow([id, new Date(), data.name || 'Anonymous Client', rating, data.text || '', 'No']);

  var webAppUrl = ScriptApp.getService().getUrl();
  var approveLink = webAppUrl + '?action=approve&id=' + encodeURIComponent(id);
  var deleteLink = webAppUrl + '?action=delete&id=' + encodeURIComponent(id);
  var stars = '\u2605'.repeat(rating) + '\u2606'.repeat(5 - rating);

  var htmlBody =
    '<div style="font-family:Arial,sans-serif; max-width:480px; line-height:1.5;">' +
    '<h2 style="margin-bottom:4px;">New review awaiting approval</h2>' +
    '<p><strong>Name:</strong> ' + (data.name || 'Anonymous Client') + '</p>' +
    '<p><strong>Rating:</strong> ' + stars + ' (' + rating + '/5)</p>' +
    '<p><strong>Review:</strong><br>' + (data.text || '') + '</p>' +
    '<p style="margin-top:20px;">' +
    '<a href="' + approveLink + '" ' +
    'style="display:inline-block; background:#bfa046; color:#15120d; padding:12px 26px; ' +
    'text-decoration:none; border-radius:4px; font-weight:bold; font-family:Arial,sans-serif; margin-right:10px;">' +
    'Approve &amp; Publish</a>' +
    '<a href="' + deleteLink + '" ' +
    'style="display:inline-block; background:#9a3b3b; color:#ffffff; padding:12px 26px; ' +
    'text-decoration:none; border-radius:4px; font-weight:bold; font-family:Arial,sans-serif;">' +
    'Delete This Review</a>' +
    '</p>' +
    '<p style="font-size:12px; color:#888; margin-top:16px;">Approve publishes it immediately. Delete removes it permanently — use this any time, whether it\\'s still pending or already live on the site. Ignoring this email leaves it hidden and unpublished.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: 'Review awaiting approval — ' + (data.name || 'Anonymous Client') + ' (' + rating + '\u2605)',
    htmlBody: htmlBody
  });

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

function handleApprove(ss, id) {
  var sheet = ss.getSheetByName('Reviews');
  if (!sheet) {
    return HtmlService.createHtmlOutput('<h2>No reviews found.</h2>');
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 6).setValue('Yes');
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial,sans-serif; padding:40px; text-align:center;">' +
        '<h2>Review approved and published</h2>' +
        '<p>It is now live on the NUVA-Gold website. Thank you!</p></div>'
      );
    }
  }
  return HtmlService.createHtmlOutput('<h2>Review not found — it may already be approved.</h2>');
}

function handleDelete(ss, id) {
  var sheet = ss.getSheetByName('Reviews');
  if (!sheet) {
    return HtmlService.createHtmlOutput('<h2>No reviews found.</h2>');
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial,sans-serif; padding:40px; text-align:center;">' +
        '<h2>Review deleted</h2>' +
        '<p>It has been permanently removed and will not appear on the website.</p></div>'
      );
    }
  }
  return HtmlService.createHtmlOutput('<h2>Review not found — it may already be deleted.</h2>');
}

function handleManage(ss) {
  var sheet = ss.getSheetByName('Reviews');
  var webAppUrl = ScriptApp.getService().getUrl();
  var rows = sheet ? sheet.getDataRange().getValues().slice(1) : [];
  rows.reverse();

  var html = '<div style="font-family:Arial,sans-serif; max-width:640px; margin:0 auto; padding:30px;">' +
    '<h2>NUVA-Gold — Manage Reviews</h2>';

  if (rows.length === 0) {
    html += '<p>No reviews submitted yet.</p>';
  }

  rows.forEach(function(row) {
    var id = row[0], name = row[2], rating = row[3], text = row[4], approved = row[5];
    var stars = '\u2605'.repeat(Number(rating) || 0) + '\u2606'.repeat(5 - (Number(rating) || 0));
    var statusLabel = approved === 'Yes' ? '<span style="color:#2e7d4f;">Published</span>' : '<span style="color:#a67c00;">Pending</span>';
    html += '<div style="border:1px solid #ddd; border-radius:8px; padding:16px; margin-bottom:14px;">' +
      '<p style="margin:0 0 6px;"><strong>' + name + '</strong> — ' + stars + ' — ' + statusLabel + '</p>' +
      '<p style="margin:0 0 12px; color:#333;">' + text + '</p>';
    if (approved !== 'Yes') {
      html += '<a href="' + webAppUrl + '?action=approve&id=' + encodeURIComponent(id) + '" ' +
        'style="background:#bfa046; color:#15120d; padding:8px 16px; text-decoration:none; border-radius:4px; font-weight:bold; margin-right:8px;">Approve</a>';
    }
    html += '<a href="' + webAppUrl + '?action=delete&id=' + encodeURIComponent(id) + '" ' +
      'style="background:#9a3b3b; color:#fff; padding:8px 16px; text-decoration:none; border-radius:4px; font-weight:bold;">Delete</a>' +
      '</div>';
  });

  html += '</div>';
  return HtmlService.createHtmlOutput(html);
}

function getApprovedReviews(ss) {
  var sheet = ss.getSheetByName('Reviews');
  if (!sheet) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var reviews = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === 'Yes') {
      reviews.push({
        timestamp: data[i][1],
        name: data[i][2],
        rating: data[i][3],
        text: data[i][4]
      });
    }
  }
  reviews.reverse();
  return ContentService.createTextOutput(JSON.stringify(reviews)).setMimeType(ContentService.MimeType.JSON);
}
```

3. Save (Ctrl/Cmd+S). Name the project `NUVA-Gold Backend`.

### Step 3: Deploy it as a Web App
1. **Deploy → New deployment**.
2. Gear icon next to "Select type" → **Web app**.
3. Settings:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. **Deploy**, then authorize when prompted (it's your own script — this is expected).
5. Copy the **Web app URL** — looks like `https://script.google.com/macros/s/AKfycb.../exec`.

### Step 4: Connect it to your website
In `index.html`, find this line (appears once, used by both forms):
```javascript
var BACKEND_ENDPOINT = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
```
Replace the placeholder with your Web app URL, save, commit, and push. Netlify redeploys automatically.

### Step 5: Test both flows
- **Contact form:** submit it → check the Sheet's "Contact" tab and your email.
- **Reviews:** submit a review on the site → check your email for the approval message → click **Approve & Publish This Review** → refresh the website. The review should now appear in the Reviews section, and the Sheet's "Reviews" tab will show "Yes" in the Approved column.

---

## How moderation works, in short
- Every review submitted goes to the "Reviews" tab with `Approved = No` and is **not shown on the site**.
- You get an email with the review and two buttons: gold **Approve & Publish**, and red **Delete This Review**.
- Approve flips that row to `Approved = Yes` and it goes live within seconds — no redeploy needed.
- Delete removes the row permanently — works whether the review is still pending or already published, so you can also remove a review after the fact if needed later.
- Ignoring the email leaves it pending and hidden — no action needed to "reject" it.

### Managing all reviews in one place
Visit `YOUR_WEB_APP_URL?action=manage` (just add `?action=manage` to the Web app URL from Step 3) to see every review ever submitted — published and pending — each with its own Approve/Delete buttons. Worth bookmarking this page for quick cleanup anytime, rather than digging through old emails.
