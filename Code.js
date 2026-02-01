const MAIN_SHEET = "Large DME Form";
const OPT_OUT_SHEET = "Opt Out";
const OPT_OUT_FORM_URL = "https://forms.gle/M2o78TFUYKptG9vG6";
const FORM_SUBMISSION_URL = "https://forms.gle/Rknq9uPDAzJGki4d7";

// Exact strings from your form
const OPT_IN_YES = "Yes - I would like to receive notifications";
const STATUS_OPTED_OUT = "Opted Out";

// Triggered on Form Submit
function onFormSubmit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const newRowValues = e.values; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET);
  const mainData = mainSheet.getDataRange().getValues();
  const headers = mainData[0];
  
  // Ensures tracking for: Opt Out, Notification Count, Initial Match Count, and Successful Match
  const colIndices = ensureTrackingColumns(mainSheet, headers);

  // --- PATH A: User submitted a new listing to Large DME Form ---
  if (sheetName === MAIN_SHEET) {
    const userEmail = newRowValues[1];
    const actionType = newRowValues[2]; // "Donate" or "Receive"
    const firstName = newRowValues[3];
    
    // Column I (8) for Donate, Column J (9) for Receive
    const itemName = (actionType === "Donate") ? newRowValues[8] : newRowValues[9]; 
    const targetAction = (actionType === "Donate") ? "Receive" : "Donate";

    // Sync with existing opt-outs before running match logic
    refreshOptOutStatus(mainSheet, colIndices);
    const updatedData = mainSheet.getDataRange().getValues();

    // Find matches for the new submitter
    const matches = getMatches(updatedData, headers, itemName, targetAction, userEmail, colIndices);
    
    // NEW: Record the number of initial matches found in the "Initial Match Count" column
    mainSheet.getRange(range.getRow(), colIndices.matchCount + 1).setValue(matches.length);

    // Send emails
    sendSubmitterEmail(userEmail, firstName, itemName, matches, headers, colIndices);
    notifyExistingSubscribers(mainSheet, updatedData, headers, newRowValues, itemName, targetAction, colIndices);
  } 

  // User submitted the Opt-Out form
  else if (sheetName === OPT_OUT_SHEET) {
    const optOutEmail = newRowValues[1];
    const optOutItem = newRowValues[3];
    const wasSuccessful = newRowValues[4];
    const partnerEmail = newRowValues[6];

    let foundInMain = false;
    
    for (let i = 1; i < mainData.length; i++) {
      const row = mainData[i];
      const rowItem = (row[2] === "Donate") ? row[8] : row[9];
      const rowEmail = row[1];
      
      // Handle the person who submitted the form
      if (normalize(rowEmail) === normalize(optOutEmail) && normalize(rowItem) === normalize(optOutItem)) {
        foundInMain = true;
        mainSheet.getRange(i + 1, colIndices.optOut + 1).setValue(STATUS_OPTED_OUT);
        mainSheet.getRange(i + 1, colIndices.success + 1).setValue(wasSuccessful);
      }
      
      // Handle the Partner Email (Column G)
      if (partnerEmail && normalize(rowEmail) === normalize(partnerEmail) && normalize(rowItem) === normalize(optOutItem)) {
        if (row[colIndices.optOut] !== STATUS_OPTED_OUT) {
          mainSheet.getRange(i + 1, colIndices.optOut + 1).setValue(STATUS_OPTED_OUT);
          mainSheet.getRange(i + 1, colIndices.success + 1).setValue("Yes");
          
          // --- NEW: Trigger notification to the Partner ---
          sendPartnerOptOutNotification(rowEmail, row[3], optOutEmail, optOutItem);
        }
      }
    }
    
    sendOptOutConfirmation(optOutEmail, optOutItem, foundInMain);
  }
}

// Sends confirmation if record found, or troubleshooting email if not.
function sendOptOutConfirmation(email, item, found) {
  let subject, body;

  if (found) {
    subject = `Confirmation: Opt-Out for ${item}`;
    body = `<p>Hello,</p>
            <p>This email confirms that we have located your record and you have successfully opted out of notifications for <b>${item}</b>.</p>
            <p>Your listing is now inactive. If you have a different item to list, you can resubmit here: <a href="${FORM_SUBMISSION_URL}">${FORM_SUBMISSION_URL}</a></p>`;
  } else {
    subject = `Action Required: Opt-Out Unsuccessful`;
    body = `<p>Hello,</p>
            <p>We received an opt-out request for <b>${item}</b>, but <b>we were unable to find a matching submission in our records.</b></p>
            <p>To stop notifications, please <a href="${OPT_OUT_FORM_URL}">submit the Opt-Out form again</a> ensuring you use the <b>exact email</b> and <b>item name</b> from your original submission.</p>`;
  }

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body + `<p>ReCARES Large DME System</p>` });
}

// Matching Logic: Finds compatible rows based on Item and Action
function getMatches(data, headers, itemName, targetAction, excludeEmail, colIndices) {
  const matches = [];
  const now = new Date();
  const lastRowIndex = data.length - 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (i === lastRowIndex) continue; // exclude current submission

    const rowAction = row[2];
    const rowItem = (rowAction === "Donate") ? row[8] : row[9];
    const ageInDays = (now - new Date(row[0])) / (1000 * 60 * 60 * 24);
    const rowOptOutStatus = row[colIndices.optOut];

    if (normalize(rowItem) === normalize(itemName) && 
        rowAction === targetAction && 
        ageInDays <= 90 && 
        rowOptOutStatus !== STATUS_OPTED_OUT) {
      matches.push(row);
    }
  }
  return matches;
}

// Notifies existing users that a new item matching theirs has been posted
function notifyExistingSubscribers(sheet, data, headers, newRow, newItemName, targetAction, colIndices) {
  let optInIdx = -1;
  for (let k = headers.length - 1; k >= 0; k--) {
    if (headers[k].toLowerCase().includes("notifications")) { optInIdx = k; break; }
  }
  
  const now = new Date();

  for (let i = 1; i < data.length - 1; i++) {
    const row = data[i];
    const rowEmail = row[1];
    const rowAction = row[2];
    const rowItem = (rowAction === "Donate") ? row[8] : row[9];
    const rowOptIn = row[optInIdx];
    const rowStatus = row[colIndices.optOut];
    const ageInDays = (now - new Date(row[0])) / (1000 * 60 * 60 * 24);

    if (normalize(rowItem) === normalize(newItemName) && 
        rowAction === targetAction && 
        rowOptIn === OPT_IN_YES && 
        rowStatus !== STATUS_OPTED_OUT && 
        ageInDays <= 90) {
      
      let userMatches = getMatches(data, headers, rowItem, newRow[2], rowEmail, colIndices);
      userMatches.push(newRow); 

      let subject = `New Match Alert: ${rowItem}`;
      let body = `<p>Hello ${row[3]},</p>
                  <p>A new match for your <b>${rowItem}</b> has been found!</p>
                  ${generateHtmlTable(userMatches, headers, newRow[0], colIndices)} 
                  <hr>
                  <p style="color: gray; font-size: 12px;">To stop these emails, fill out the <a href="${OPT_OUT_FORM_URL}">Opt-Out Form</a>.</p>`;
      
      MailApp.sendEmail({ to: rowEmail, subject: subject, htmlBody: body });

      let currentCount = parseInt(row[colIndices.count]) || 0;
      sheet.getRange(i + 1, colIndices.count + 1).setValue(currentCount + 1);
    }
  }
}

// Formats the HTML table for emails, including privacy logic
function generateHtmlTable(rows, headers, highlightTimestamp = null, colIndices) {
  let optInIdx = -1;
  for (let k = headers.length - 1; k >= 0; k--) {
    if (headers[k].toLowerCase().includes("notifications")) { optInIdx = k; break; }
  }
  const phoneIdx = 6; 
  const permitIdx = 7; 
  
  let table = `<table border="1" style="border-collapse: collapse; width: 100%; font-family: sans-serif; font-size: 13px;">
                <tr style="background-color: #4A90E2; color: white;">
                  <th style="padding: 8px;">Posted</th><th style="padding: 8px;">Item</th><th style="padding: 8px;">Name</th>
                  <th style="padding: 8px;">City</th><th style="padding: 8px;">Contact</th><th style="padding: 8px;">Details</th>
                </tr>`;
  
  rows.forEach(row => {
    const isNew = highlightTimestamp && row[0].toString() === highlightTimestamp.toString();
    const bgColor = isNew ? "#FFFFCC" : "#FFFFFF"; 
    const rowItemName = row[2] === "Donate" ? row[8] : row[9];
    const d = new Date(row[0]);
    const formattedDate = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear().toString().slice(-2);

    let contactInfo = `<a href="mailto:${row[1]}">${row[1]}</a>`;
    const permissionValue = (row[permitIdx] || "").toString().toLowerCase();
    if (permissionValue.includes("yes") && row[phoneIdx]) {
      contactInfo += `, ${row[phoneIdx]}`;
    }
    
    let detailParts = [];
    for (let i = 10; i < row.length; i++) {
      if (i === optInIdx || i === colIndices.optOut || i === colIndices.count || i === colIndices.matchCount || i === colIndices.success) continue;
      if (row[i] && row[i].toString().trim() !== "" && row[i].toString().toLowerCase() !== "agree") {
        detailParts.push(row[i]);
      }
    }

    table += `<tr style="background-color: ${bgColor};">
                <td style="padding: 8px; white-space: nowrap;">${formattedDate}</td>
                <td style="padding: 8px;"><b>${rowItemName}</b></td>
                <td style="padding: 8px;">${row[3]}</td>
                <td style="padding: 8px;">${row[5]}</td>
                <td style="padding: 8px;">${contactInfo}</td>
                <td style="padding: 8px;">${detailParts.join(" • ")}</td>
              </tr>`;
  });
  return table + `</table>`;
}

// Handles column creation and returns current column indices
function ensureTrackingColumns(sheet, headers) {
  const columnsToAdd = [
    { name: "Opt Out Status", key: "optOut" },
    { name: "Notification Count", key: "count" },
    { name: "Initial Match Count", key: "matchCount" },
    { name: "Successful Match", key: "success" }
  ];

  let indices = {};
  columnsToAdd.forEach(col => {
    let idx = headers.indexOf(col.name);
    if (idx === -1) {
      sheet.getRange(1, headers.length + 1).setValue(col.name);
      idx = headers.length;
      headers.push(col.name);
    }
    indices[col.key] = idx;
  });
  return indices;
}

// Daily function to send 83-day warnings
function checkExpirations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAIN_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndices = ensureTrackingColumns(sheet, headers);
  
  let optInIdx = -1;
  for (let k = headers.length - 1; k >= 0; k--) {
    if (headers[k].toLowerCase().includes("notifications")) { optInIdx = k; break; }
  }

  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = new Date(row[0]);
    const ageInDays = Math.floor((now - timestamp) / (1000 * 60 * 60 * 24));

    if (ageInDays === 83) {
      const rowStatus = row[colIndices.optOut];
      if (row[optInIdx] === OPT_IN_YES && rowStatus !== STATUS_OPTED_OUT) {
        const rowItem = (row[2] === "Donate") ? row[8] : row[9];
        let subject = `Action Required: Your ${rowItem} listing expires in 7 days`;
        let body = `<p>Hello ${row[3]},</p>
                    <p>Your listing for <b>${rowItem}</b> expires in 7 days.</p>
                    <p>To stay in the matching system, please resubmit here: <a href="${FORM_SUBMISSION_URL}">${FORM_SUBMISSION_URL}</a></p>`;
        MailApp.sendEmail({ to: row[1], subject: subject, htmlBody: body });
      }
    }
  }
}

// Syncs the Main Sheet with all entries from the Opt Out sheet
function refreshOptOutStatus(sheet, colIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const optOutSheet = ss.getSheetByName(OPT_OUT_SHEET);
  if (!optOutSheet) return;
  const mainData = sheet.getDataRange().getValues();
  const optOutData = optOutSheet.getDataRange().getValues();

  for (let i = 1; i < mainData.length; i++) {
    const mRow = mainData[i];
    const mItem = (mRow[2] === "Donate") ? mRow[8] : mRow[9];
    const mEmail = normalize(mRow[1]);
    
    for (let j = 1; j < optOutData.length; j++) {
      const oRow = optOutData[j];
      const oSubmitterEmail = normalize(oRow[1]);
      const oItem = normalize(oRow[3]);
      const oPartnerEmail = normalize(oRow[6]); // Partner Email from Col G

      if (normalize(mItem) === oItem) {
        // Match found if main row is the submitter OR the partner
        if (mEmail === oSubmitterEmail || (oPartnerEmail !== "" && mEmail === oPartnerEmail)) {
          sheet.getRange(i + 1, colIndices.optOut + 1).setValue(STATUS_OPTED_OUT);
          
          // If the main row is the partner, we force success to "Yes"
          const successVal = (mEmail === oPartnerEmail) ? "Yes" : oRow[4];
          sheet.getRange(i + 1, colIndices.success + 1).setValue(successVal);
        }
      }
    }
  }
}

function normalize(str) {
  if (!str) return "";
  return str.toString().toLowerCase().trim().replace(/s$/, ""); 
}

function sendSubmitterEmail(email, name, item, matches, headers, colIndices) {
  let subject = `Matches found for your ${item}`;
  let body = `<p>Hello ${name},</p>`;
  if (matches.length > 0) {
    body += `<p>We found matches for your <b>${item}</b>:</p>${generateHtmlTable(matches, headers, null, colIndices)}`;
  } else {
    body += `<p>No current matches for <b>${item}</b>. We will notify you if a new match appears.</p>`;
  }
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body + `<p>ReCARES Large DME System</p>` });
}