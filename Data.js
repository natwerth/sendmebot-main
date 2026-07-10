// --- Senders sheet helpers: sender table A:C, image table E:G ---

function getSenderProfile_(ss, senderEmail) {
  const sheet = ss.getSheetByName("Senders");

  if (!sheet) {
    return {
      name: "",
      email: senderEmail,
      signatureText: "",
      signatureRichText: null
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      name: "",
      email: senderEmail,
      signatureText: "",
      signatureRichText: null
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const name = String(values[i][0] || "").trim();
    const email = String(values[i][1] || "").trim();

    if (normalize_(email) === normalize_(senderEmail)) {
      const signatureCell = sheet.getRange(rowNumber, 3);

      return {
        name: name,
        email: email,
        signatureText: signatureCell.getDisplayValue(),
        signatureRichText: signatureCell.getRichTextValue()
      };
    }
  }

  return {
    name: "",
    email: senderEmail,
    signatureText: "",
    signatureRichText: null
  };
}


function getSenderEmails_(sendersSheet) {
  const lastRow = sendersSheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sendersSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const emails = [];

  values.forEach(row => {
    const email = String(row[0] || "").trim();

    if (email && emails.indexOf(email) === -1) {
      emails.push(email);
    }
  });

  return emails;
}

function saveSenderProfile(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Senders");

  if (!sheet) {
    sheet = ss.insertSheet("Senders");
    sheet.getRange(1, 1, 1, 3).setValues([[
      "Sender Name",
      "Sender Email",
      "Sender Signature"
    ]]);
  }

  const name = String(formData.name || "").trim();
  const email = String(formData.email || "").trim();
  const signature = String(formData.signature || "").trim();

  if (!name) throw new Error("Sender name is required.");
  if (!email) throw new Error("Sender email is required.");

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("Enter a valid sender email address.");
  }

  const lastRow = sheet.getLastRow();
  let targetRow = null;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    for (let i = 0; i < values.length; i++) {
      const existingEmail = String(values[i][1] || "").trim();

      if (normalize_(existingEmail) === normalize_(email)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
  }

  sheet.getRange(targetRow, 1).setValue(name);
  sheet.getRange(targetRow, 2).setValue(email);

  const signatureRange = sheet.getRange(targetRow, 3);

  if (signature) {
    signatureRange.setRichTextValue(buildRichTextValueFromTemplateHtml_(signature));
  } else {
    signatureRange.setValue("");
  }

  sheet.getRange(targetRow, 1, 1, 3).setWrap(true);

  refreshFormCache();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Sender saved: " + email,
    "SendMeBot",
    5
  );

  return {
    title: "Sender saved",
    message: name + " <" + email + ">"
  };
}

// --- Image helpers ---

function getImageAssets_(ss) {
  const sheet = ss.getSheetByName("Senders");
  const assets = {};

  if (!sheet) return assets;

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return assets;

  const values = sheet.getRange(2, 5, lastRow - 1, 3).getValues();

  values.forEach(row => {
    const name = String(row[0] || "").trim();
    const link = String(row[1] || "").trim();
    const width = String(row[2] || "").trim();

    if (!name || !link) return;

    assets[normalizeTemplateKey_(name)] = {
      name: name,
      link: link,
      width: width
    };
  });

  return assets;
}


function getImageHtmlForAsset_(asset, inlineImages) {
  if (!asset || !asset.link) return "";

  const cid = getAssetCid_(asset.name);

  if (!inlineImages[cid]) {
    const blob = getDriveImageBlob_(asset.link);
    inlineImages[cid] = blob;
  }

  const width = asset.width ? parseInt(asset.width, 10) : "";

  let style = "height:auto;display:block;";

  if (width && width > 0) {
    style += "max-width:" + width + "px;";
  } else {
    style += "max-width:600px;";
  }

  return '<img src="cid:' + cid + '" alt="' + escapeHtml_(asset.name) + '" style="' + style + '">';
}


function getAssetCid_(name) {
  return "asset_" + normalizeTemplateKey_(name).replace(/[^a-z0-9]+/g, "_");
}


function getDriveImageBlob_(driveLinkOrId) {
  const fileId = extractDriveFileId_(driveLinkOrId);

  if (!fileId) {
    throw new Error("Could not extract Drive file ID from image link.");
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();

  return blob.setName(file.getName());
}


function extractDriveFileId_(value) {
  const raw = String(value || "").trim();

  if (!raw) return "";

  let match = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = raw.match(/\/uc\?id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  if (/^[a-zA-Z0-9_-]{20,}$/.test(raw)) return raw;

  return "";
}

function saveImageAsset(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Senders");

  if (!sheet) {
    sheet = ss.insertSheet("Senders");
  }

  const imageName = String(formData.name || "").trim();
  const imageLink = String(formData.link || "").trim();
  const width = String(formData.width || "").trim();

  if (!imageName) throw new Error("Image name is required.");
  if (!imageLink) throw new Error("Drive link or file ID is required.");

  const fileId = extractDriveFileId_(imageLink);

  if (!fileId) {
    throw new Error("Could not identify a valid Google Drive file ID from that link.");
  }

  try {
    DriveApp.getFileById(fileId);
  } catch (err) {
    throw new Error("Could not access that Drive file. Check the link and sharing permissions.");
  }

  if (width && !/^\d+$/.test(width)) {
    throw new Error("Width must be a whole number, like 320.");
  }

  // Ensure image table headers in E:G.
  sheet.getRange(1, 5, 1, 3).setValues([[
    "Image Name",
    "Drive Link",
    "Width"
  ]]);

  const lastRow = sheet.getLastRow();
  let targetRow = null;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 5, lastRow - 1, 3).getValues();

    for (let i = 0; i < values.length; i++) {
      const existingName = String(values[i][0] || "").trim();

      if (normalizeTemplateKey_(existingName) === normalizeTemplateKey_(imageName)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (!targetRow) {
    targetRow = Math.max(sheet.getLastRow() + 1, 2);
  }

  sheet.getRange(targetRow, 5).setValue(imageName);
  sheet.getRange(targetRow, 6).setValue(imageLink);
  sheet.getRange(targetRow, 7).setValue(width || "320");

  sheet.getRange(targetRow, 5, 1, 3).setWrap(true);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Image saved: " + imageName,
    "SendMeBot",
    5
  );

  return {
    title: "Image saved",
    message: imageName
  };
}

// --- Selection helpers ---

function getSelectedRows_(sheet, headers) {
  const selectCol = headers["select"];
  if (!selectCol) throw new Error("Missing Select column.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet
    .getRange(2, selectCol, lastRow - 1, 1)
    .getValues()
    .map((value, index) => value[0] === true ? index + 2 : null)
    .filter(Boolean);
}


function clearSelectedRows_(sheet, headers, rows) {
  const selectCol = headers["select"];
  if (!selectCol) return;

  rows.forEach(row => sheet.getRange(row, selectCol).setValue(false));
}


function getSelectedRowCountOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getTrackerSheet_();

  if (!trackerSheet) throw new Error("Missing Tracker sheet.");

  const headers = getHeaders_(trackerSheet);
  const selectedRows = getSelectedRows_(trackerSheet, headers);

  return {
    selectedRowCount: selectedRows.length
  };
}


// --- Status stamping ---

function stampTemplateColumn_(sheet, row, headers, key) {
  const col = getTemplateStatusColumn_(headers, key, true);
  const formatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d");

  if (col) sheet.getRange(row, col).setValue("Sent on " + formatted);
}


function stampTemplateFailure_(sheet, row, headers, key) {
  const col = getTemplateStatusColumn_(headers, key, true);
  if (col) sheet.getRange(row, col).setValue("Error: Not Sent");
}


function setStatus_(sheet, row, statusCol, value) {
  if (!statusCol) return;

  sheet.getRange(row, statusCol).setValue(value);
  SpreadsheetApp.flush();
}


// --- Logging ---

function logSentEmail_(ss, logData) {
  const logSheet = ss.getSheetByName("Sent");
  if (!logSheet) throw new Error("Missing Sent sheet.");

  const headers = ensureLogHeaders_(logSheet);

  const timestamp = Object.prototype.hasOwnProperty.call(logData, "timestamp")
    ? logData.timestamp
    : Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "M/d/yyyy h:mm:ss a"
    );

  logSheet.insertRowBefore(2);

  const rowData = {
    "Timestamp": timestamp,
    "Status": logData.status || "",
    "Scheduled For": logData.scheduledFor || "",
    "Processed At": logData.processedAt || "",
    "Message": logData.message || "",
    "Name": logData.name || "",
    "Recipient": logData.email || "",
    "Sender": logData.sender || "",
    "CC": logData.cc || "",
    "BCC": logData.bcc || "",
    "Template": logData.template || "",
    "Subject": logData.subject || "",
    "Email Body": logData.body || "",
    "Attachments": logData.attachments || "",
    "Log Note": logData.logNote || ""
  };

  const output = new Array(logSheet.getLastColumn()).fill("");

  Object.keys(rowData).forEach(header => {
    const col = headers[normalize_(header)];
    if (col) output[col - 1] = rowData[header];
  });

  logSheet.getRange(2, 1, 1, output.length).setValues([output]);

  return 2;
}


function ensureLogHeaders_(logSheet) {
  const expectedHeaders = [
    "Timestamp",
    "Status",
    "Scheduled For",
    "Processed At",
    "Message",
    "Name",
    "Recipient",
    "Sender",
    "CC",
    "BCC",
    "Template",
    "Subject",
    "Email Body",
    "Attachments",
    "Log Note"
  ];

  const lastCol = Math.max(logSheet.getLastColumn(), expectedHeaders.length);
  const existingValues = logSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const existingHeaders = {};

  existingValues.forEach((header, index) => {
    const key = normalize_(header);
    if (key) existingHeaders[key] = index + 1;
  });

  // Rename old Email header to Recipient if present.
  if (existingHeaders["email"] && !existingHeaders["recipient"]) {
    logSheet.getRange(1, existingHeaders["email"]).setValue("Recipient");
    existingHeaders["recipient"] = existingHeaders["email"];
    delete existingHeaders["email"];
  }

  let nextCol = logSheet.getLastColumn();

  expectedHeaders.forEach(header => {
    const key = normalize_(header);

    if (!existingHeaders[key]) {
      nextCol++;
      logSheet.getRange(1, nextCol).setValue(header);
      existingHeaders[key] = nextCol;
    }
  });

  return getHeaders_(logSheet);
}

// --- Row data / generic template filling ---

function getRowData_(sheet, row, headers) {
  const data = {};

  Object.keys(headers).forEach(headerName => {
    data[toDisplayHeader_(headerName)] = sheet
      .getRange(row, headers[headerName])
      .getDisplayValue();
  });

  return data;
}


function fillTemplate_(text, data) {
  return String(text || "").replace(/\{\{(.*?)\}\}/g, (_, key) => {
    return data[key.trim()] || "";
  });
}


// --- Generic sheet helpers ---

function getCellValue_(sheet, row, headers, headerName) {
  const col = headers[normalize_(headerName)];
  return col ? sheet.getRange(row, col).getValue() : "";
}


function getHeaders_(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = {};

  raw.forEach((header, index) => {
    const key = normalize_(header);
    if (key) headers[key] = index + 1;
  });

  return headers;
}


// --- HTML / text helpers ---

function textToHtml_(text) {
  return escapeHtml_(String(text || "")).replace(/\n/g, "<br>");
}


function escapeHtml_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


// --- String helpers ---

function toDisplayHeader_(headerName) {
  return headerName
    .split(" ")
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}


function toTitleCase_(value) {
  return String(value || "")
    .split(" ")
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}


function normalize_(value) {
  return String(value || "").trim().toLowerCase();
}


function normalizeTemplateKey_(value) {
  return normalize_(value)
    .replace(/[_\s]+/g, " ")
    .trim();
}

function getDefaultEmailWrapperStyle_() {
  return "font-family: Arial, sans-serif; font-size: 15px; line-height: 1.4;";
}


function shouldPreserveRichTextFontSize_(fontSize) {
  const size = Number(fontSize);

  if (!size) return false;

  // Treat normal Google Sheets defaults as "no explicit override."
  // This prevents template/sig text from shrinking just because Sheets defaulted to 10pt/11pt.
  if (size === 10 || size === 11) return false;

  // Preserve deliberate-looking manual sizes.
  return size >= 6 && size <= 72;
}
