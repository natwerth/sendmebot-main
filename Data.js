// --- Senders sheet helpers: sender table A:C, separator D, image table E:G ---

function getAuthenticatedUserEmail_(runtime) {
  const session = runtime && runtime.session ? runtime.session : Session;
  const email = String(session.getEffectiveUser().getEmail() || "").trim();

  if (!email) {
    throw new Error(
      "Your authenticated Google account could not be determined. " +
      "Reload the spreadsheet and authorize SendMeBot before continuing."
    );
  }

  return email;
}

function getSenderTableColumns_(sheet, allowInitialize) {
  const rawHeaders = sheet.getRange(1, 1, 1, 3).getValues()[0];
  const allBlank = rawHeaders.every(value => !String(value || "").trim());

  if (allBlank && allowInitialize) {
    sheet.getRange(1, 1, 1, 3).setValues([["Name", "Email", "Signature"]]);
    return { name: 1, email: 2, signature: 3 };
  }
  if (allBlank) return null;

  const columns = {};
  rawHeaders.forEach((header, index) => {
    const key = normalize_(header).replace(/^sender\s+/, "");
    if (key === "name" || key === "email" || key === "signature") {
      columns[key] = index + 1;
    }
  });

  if (!columns.name || !columns.email || !columns.signature) {
    throw new Error(
      "The Senders sender table must use Name, Email, and Signature in columns A:C."
    );
  }

  return columns;
}

function getAuthorizedSenderOptions_(ss) {
  const sheet = ss.getSheetByName("Senders");
  if (!sheet) return [];

  const columns = getSenderTableColumns_(sheet, false);
  if (!columns || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .map(rowValues => ({
      name: String(rowValues[columns.name - 1] || "").trim(),
      email: String(rowValues[columns.email - 1] || "").trim()
    }))
    .filter(sender => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(sender.email));
}

function getSenderMatches_(sheet, senderEmail) {
  const columns = getSenderTableColumns_(sheet, false);
  if (!columns) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const matches = [];

  values.forEach((rowValues, index) => {
    const email = String(rowValues[columns.email - 1] || "").trim();
    if (normalize_(email) !== normalize_(senderEmail)) return;

    const row = index + 2;
    const signatureCell = sheet.getRange(row, columns.signature);
    matches.push({
      row: row,
      name: String(rowValues[columns.name - 1] || "").trim(),
      email: email,
      signatureText: signatureCell.getDisplayValue(),
      signatureRichText: signatureCell.getRichTextValue()
    });
  });

  return matches;
}

function getSenderStateForEmail_(ss, senderEmail) {
  const sheet = ss.getSheetByName("Senders");
  if (!sheet) {
    return { status: "missing", email: senderEmail, record: null, duplicateCount: 0 };
  }

  const matches = getSenderMatches_(sheet, senderEmail);
  if (matches.length > 1) {
    return {
      status: "duplicate",
      email: senderEmail,
      record: null,
      duplicateCount: matches.length
    };
  }

  if (!matches.length) {
    return { status: "missing", email: senderEmail, record: null, duplicateCount: 0 };
  }

  if (!matches[0].name) {
    return { status: "invalid", email: senderEmail, record: matches[0], duplicateCount: 0 };
  }

  return { status: "ready", email: senderEmail, record: matches[0], duplicateCount: 0 };
}

function getSenderProfile_(ss, senderEmail) {
  const state = getSenderStateForEmail_(ss, senderEmail);
  if (state.status === "duplicate") {
    throw new Error(
      "Duplicate sender profiles exist for " + senderEmail + ". Resolve them before sending."
    );
  }

  if (state.status === "ready") return state.record;

  return {
    name: "",
    email: senderEmail,
    signatureText: "",
    signatureRichText: null
  };
}

function getSenderSignatureEditorHtml_(profile) {
  const richText = profile && profile.signatureRichText;
  if (!richText || typeof richText.getRuns !== "function") {
    return escapeHtml_(String((profile && profile.signatureText) || "")).replace(/\n/g, "<br>");
  }

  return richText.getRuns().map(run => {
    let html = escapeHtml_(String(run.getText() || "")).replace(/\n/g, "<br>");
    const style = run.getTextStyle ? run.getTextStyle() : null;
    let css = "";

    if (style) {
      if (style.isBold && style.isBold()) html = "<strong>" + html + "</strong>";
      const color = style.getForegroundColor ? style.getForegroundColor() : "";
      const size = style.getFontSize ? style.getFontSize() : "";
      if (color) css += "color:" + color + ";";
      if (size) css += "font-size:" + size + "pt;";
    }

    if (css) html = '<span style="' + css + '">' + html + "</span>";

    const link = run.getLinkUrl ? String(run.getLinkUrl() || "") : "";
    if (/^(https?:\/\/|mailto:)/i.test(link)) {
      html = '<a href="' + escapeHtml_(link) + '">' + html + "</a>";
    }

    return html;
  }).join("");
}

function getCurrentUserSenderState_(ss, runtime) {
  const email = runtime && runtime.authenticatedEmail
    ? String(runtime.authenticatedEmail).trim()
    : getAuthenticatedUserEmail_(runtime);
  return getSenderStateForEmail_(ss, email);
}

function findAvailableSenderRow_(sheet, columns) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow < 2) return 2;

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < values.length; i++) {
    const hasSenderData = [columns.name, columns.email, columns.signature]
      .some(col => String(values[i][col - 1] || "").trim());
    if (!hasSenderData) return i + 2;
  }

  return lastRow + 1;
}

function saveSenderProfile(formData) {
  return saveSenderProfile_(formData);
}

function saveSenderProfile_(formData, runtime) {
  const options = runtime || {};
  const ss = options.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const authenticatedEmail = options.authenticatedEmail || getAuthenticatedUserEmail_(options);
  const submittedEmail = String((formData && formData.email) || "").trim();
  const name = String((formData && formData.name) || "").trim();
  const signature = String((formData && formData.signature) || "").trim();

  if (submittedEmail && normalize_(submittedEmail) !== normalize_(authenticatedEmail)) {
    throw new Error("Sender email must match your authenticated Google account.");
  }
  if (!name) throw new Error("Sender name is required.");

  const lock = options.lock || LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    let sheet = ss.getSheetByName("Senders");
    if (!sheet) sheet = ss.insertSheet("Senders");

    const columns = getSenderTableColumns_(sheet, true);
    const matches = getSenderMatches_(sheet, authenticatedEmail);

    if (matches.length > 1) {
      throw new Error(
        "Duplicate sender profiles exist for " + authenticatedEmail +
        ". Remove the duplicate rows before continuing."
      );
    }

    const targetRow = matches.length
      ? matches[0].row
      : findAvailableSenderRow_(sheet, columns);

    sheet.getRange(targetRow, columns.name).setValue(name);
    sheet.getRange(targetRow, columns.email).setValue(authenticatedEmail);

    const signatureRange = sheet.getRange(targetRow, columns.signature);
    if (signature) {
      const buildRichText = options.buildRichText || buildRichTextValueFromTemplateHtml_;
      signatureRange.setRichTextValue(buildRichText(signature));
    } else {
      signatureRange.setValue("");
    }

    sheet.getRange(targetRow, 1, 1, 3).setWrap(true);

    const savedState = getSenderStateForEmail_(ss, authenticatedEmail);
    if (savedState.status !== "ready") {
      throw new Error("Sender profile could not be verified after saving.");
    }

    if (!options.suppressToast) {
      ss.toast("Sender saved: " + authenticatedEmail, "SendMeBot", 5);
    }

    return {
      title: "Sender saved",
      message: name + " <" + authenticatedEmail + ">",
      senderState: {
        status: "ready",
        email: authenticatedEmail,
        duplicateCount: 0,
        record: {
          name: savedState.record.name,
          email: authenticatedEmail,
          signature: savedState.record.signatureText || ""
        }
      }
    };
  } finally {
    lock.releaseLock();
  }
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


function getImageHtmlForAsset_(asset, inlineImages, assetContext) {
  if (!asset || !asset.link) return "";

  const cid = getAssetCid_(asset.name);

  if (!inlineImages[cid]) {
    const blob = getDriveImageBlob_(asset.link, asset, assetContext);
    if (!blob) return "";
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


function getImageAssetKey_(asset) {
  const item = asset || {};
  const fileId = extractDriveFileId_(item.link);
  return "image:" + normalizeTemplateKey_(item.name) + ":" + (fileId || String(item.link || ""));
}


function getDriveImageBlob_(driveLinkOrId, asset, assetContext) {
  const fileId = extractDriveFileId_(driveLinkOrId);
  const item = asset || { name: "Image", link: driveLinkOrId };
  const key = getImageAssetKey_(item);

  if (!fileId) {
    if (shouldOmitAssetFailure_(assetContext, key)) {
      recordOmittedAsset_(assetContext, {
        key: key,
        kind: "image",
        label: item.name || "Image",
        reason: "The image link does not contain a valid Drive file ID."
      });
      return null;
    }
    throw new Error(
      'Could not access image "' + (item.name || "Image") + '" because its Drive link is invalid.'
    );
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

    return blob.setName(file.getName());
  } catch (err) {
    if (shouldOmitAssetFailure_(assetContext, key)) {
      recordOmittedAsset_(assetContext, {
        key: key,
        kind: "image",
        label: item.name || "Image",
        reason: "The Drive image is unavailable to the sending account."
      });
      return null;
    }
    throw new Error(
      'Could not access image "' + (item.name || "Image") +
      '". Check its Drive link and sharing permissions.'
    );
  }
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


function clearSelectedRow_(sheet, headers, row) {
  const selectCol = headers["select"];
  if (!selectCol || !row) return;

  sheet.getRange(row, selectCol).setValue(false);
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

function stampTemplateColumn_(sheet, row, headers, key, trackingEnabled) {
  if (trackingEnabled === false) return;
  const col = getTemplateStatusColumn_(headers, key);
  const formatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d");

  if (col) sheet.getRange(row, col).setValue("Sent on " + formatted);
}


function stampTemplateFailure_(sheet, row, headers, key, trackingEnabled) {
  if (trackingEnabled === false) return;
  const col = getTemplateStatusColumn_(headers, key);
  if (col) sheet.getRange(row, col).setValue("Error: Not Sent");
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
    "Record ID": logData.recordId || "",
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
    "Record ID",
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

  // Rename the legacy Sent identity header in place without disturbing data.
  if (existingHeaders["name"] && !existingHeaders["record id"]) {
    logSheet.getRange(1, existingHeaders["name"]).setValue("Record ID");
    existingHeaders["record id"] = existingHeaders["name"];
    delete existingHeaders["name"];
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
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};
  const raw = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
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
