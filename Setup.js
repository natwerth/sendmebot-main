const SENDMEBOT_INTERNAL_SHEET = "_SendMeBot";
const SENDMEBOT_SETUP_VERSION_PROPERTY = "SENDMEBOT_SETUP_VERSION";
const SENDMEBOT_SETUP_VERSION = "1";

const SENDMEBOT_RESERVED_SCHEMAS = {
  Tracker: ["Select", "Name", "Email"],
  Templates: ["Name", "Subject", "Body", "Attachment Link"],
  Senders: ["Name", "Email", "Signature", "", "Image Name", "Drive Link", "Width"],
  Sent: [
    "Timestamp", "Status", "Scheduled For", "Processed At", "Message", "Record ID",
    "Recipient", "Sender", "CC", "BCC", "Template", "Subject", "Email Body",
    "Attachments", "Log Note"
  ]
};


function initializeNewSendMeBotWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = ensureSendMeBotWorkbook_(ss, { createTracker: true, seedNeutralConfig: true });
  ensureSendMeBotInstallationRegistry_(ss);
  return result;
}


function ensureSendMeBotWorkbook_(ss, options) {
  const settings = options || {};
  const created = [];
  const repaired = [];
  const requiredNames = settings.createTracker === false
    ? ["Templates", "Senders", "Sent"]
    : ["Tracker", "Templates", "Senders", "Sent"];

  requiredNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      created.push(name);
    }
    if (ensureSendMeBotHeaders_(sheet, SENDMEBOT_RESERVED_SCHEMAS[name])) repaired.push(name);
    styleSendMeBotSheet_(sheet, name);
  });

  const tracker = ss.getSheetByName("Tracker");
  if (tracker) ensureSelectCheckboxes_(tracker);

  if (settings.seedNeutralConfig) {
    PropertiesService.getDocumentProperties().setProperties({
      [TRACKER_SHEET_PROPERTY]: "Tracker",
      [RECORD_ID_HEADER_PROPERTY]: "Name",
      [SENDMEBOT_SETUP_VERSION_PROPERTY]: SENDMEBOT_SETUP_VERSION
    });
  }

  return { created: created, repaired: repaired };
}


function ensureSendMeBotHeaders_(sheet, expected) {
  const headers = expected || [];
  if (!headers.length) return false;
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0];
  let changed = false;

  headers.forEach((header, index) => {
    if (!header) return;
    if (!String(existing[index] || "").trim()) {
      sheet.getRange(1, index + 1).setValue(header);
      changed = true;
    }
  });

  return changed;
}


function styleSendMeBotSheet_(sheet, name) {
  const schema = SENDMEBOT_RESERVED_SCHEMAS[name] || [];
  if (!schema.length) return;
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, schema.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#356854")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setRowHeight(1, 24);

  schema.forEach((header, index) => {
    if (!header) return;
    const width = header === "Body" || header === "Email Body" || header === "Message"
      ? 340
      : header === "Drive Link" || header === "Attachment Link" || header === "Attachments"
        ? 260
        : header === "Email" || header === "Recipient" || header === "Sender"
          ? 220
          : header === "Select"
            ? 70
            : 150;
    sheet.setColumnWidth(index + 1, width);
  });
}


function ensureSelectCheckboxes_(sheet) {
  const headers = getHeaders_(sheet);
  const selectCol = headers["select"];
  if (!selectCol) return false;
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, selectCol, rowCount, 1).insertCheckboxes();
  return true;
}


function addSelectColumnToTracker_(sheet) {
  const headers = getHeaders_(sheet);
  if (headers["select"]) {
    ensureSelectCheckboxes_(sheet);
    return headers["select"];
  }
  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1).setValue("Select");
  ensureSelectCheckboxes_(sheet);
  return 1;
}


function ensureSendMeBotInstallationRegistry_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SENDMEBOT_INTERNAL_SHEET);
  if (!sheet) sheet = ss.insertSheet(SENDMEBOT_INTERNAL_SHEET);

  const existing = {};
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues().forEach(row => {
      const key = String(row[0] || "").trim();
      if (key) existing[key] = String(row[1] || "");
    });
  }

  const values = [
    ["formatVersion", SENDMEBOT_INSTALL_FORMAT_VERSION],
    ["appId", SENDMEBOT_APP_ID],
    ["appVersion", SENDMEBOT_APP_VERSION],
    ["scriptId", ScriptApp.getScriptId()],
    ["spreadsheetId", ss.getId()],
    ["registeredAt", existing.registeredAt || new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, 2).setValues(values);
  sheet.hideSheet();
  return values.reduce((result, row) => {
    result[row[0]] = row[1];
    return result;
  }, {});
}


function getSendMeBotInstallationRegistry_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SENDMEBOT_INTERNAL_SHEET);
  if (!sheet) return null;
  const result = {};
  sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues().forEach(row => {
    const key = String(row[0] || "").trim();
    if (key) result[key] = String(row[1] || "");
  });
  return result;
}
