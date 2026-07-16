function getSendMeBotWalkthroughContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = ensureSendMeBotInstallationRegistry_(ss);
  const templateSheet = ss.getSheetByName("Templates");
  const templates = templateSheet ? getWalkthroughValidTemplateKeys_(templateSheet) : [];
  const config = getSendMeBotConfig_();
  let authenticatedEmail = "";
  let senderState = { status: "blocked", record: null, duplicateCount: 0 };
  let senderError = "";

  try {
    authenticatedEmail = getAuthenticatedUserEmail_();
    senderState = toSenderBootstrapState_(
      getSenderStateForEmail_(ss, authenticatedEmail),
      true
    );
  } catch (err) {
    senderError = err.message || String(err);
  }

  const sheets = ss.getSheets()
    .filter(sheet => sheet.getName() !== SENDMEBOT_INTERNAL_SHEET)
    .map(sheet => getWalkthroughSheetContext_(sheet, templates));
  const suggested = sheets.some(sheet => sheet.name === registry.suggestedTrackerSheet)
    ? registry.suggestedTrackerSheet
    : sheets.some(sheet => sheet.name === config.trackerSheetName)
      ? config.trackerSheetName
      : sheets.length ? sheets[0].name : "";

  return {
    brand: getSendMeBotClientBrand_(),
    registry: {
      onboardingState: registry.onboardingState || "",
      installMode: registry.installMode || "",
      sourceSpreadsheetName: registry.sourceSpreadsheetName || "",
      importedSheets: parseWalkthroughJsonArray_(registry.importedSheets),
      suggestedTrackerSheet: suggested,
      testEmailSentAt: registry.testEmailSentAt || ""
    },
    config: config,
    sheets: sheets,
    authenticatedEmail: authenticatedEmail,
    senderState: senderState,
    senderError: senderError,
    templates: templates,
    readiness: getWalkthroughReadiness_(ss, config, authenticatedEmail, senderState, templates)
  };
}


function parseWalkthroughJsonArray_(value) {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}


function getWalkthroughSheetContext_(sheet, templates) {
  const rawHeaders = getWalkthroughRawHeaders_(sheet);
  const headers = getHeaders_(sheet);
  const recipientFields = getRecipientFieldsForSendForm_(sheet, headers, templates || []);
  return {
    name: sheet.getName(),
    headers: rawHeaders.filter(Boolean),
    hasSelectColumn: !!headers.select,
    recipientFields: recipientFields
  };
}


function getWalkthroughRawHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(value => String(value || "").trim());
}


function getWalkthroughReadiness_(ss, config, authenticatedEmail, senderState, templates) {
  const tracker = ss.getSheetByName(config.trackerSheetName);
  const headers = tracker ? getHeaders_(tracker) : {};
  const recipientFields = tracker
    ? getRecipientFieldsForSendForm_(tracker, headers, templates || [])
    : [];
  return {
    tracker: !!(
      tracker && headers.select && headers[normalize_(config.recordIdHeader)]
    ),
    recipients: recipientFields.length > 0,
    sender: !!(
      authenticatedEmail && senderState && senderState.status === "ready"
    ),
    template: Array.isArray(templates) && templates.length > 0
  };
}


function getWalkthroughValidTemplateKeys_(templateSheet) {
  if (!templateSheet) return [];
  const headers = getHeaders_(templateSheet);
  const nameColumn = getTemplateNameCol_(headers);
  const subjectColumn = headers.subject;
  const bodyColumn = headers.body;
  if (!nameColumn || !subjectColumn || !bodyColumn || templateSheet.getLastRow() < 2) return [];

  return templateSheet
    .getRange(2, 1, templateSheet.getLastRow() - 1, templateSheet.getLastColumn())
    .getDisplayValues()
    .filter(row =>
      String(row[nameColumn - 1] || "").trim() &&
      String(row[subjectColumn - 1] || "").trim() &&
      String(row[bodyColumn - 1] || "").trim()
    )
    .map(row => String(row[nameColumn - 1] || "").trim());
}


function saveWalkthroughTrackerSetup(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = formData || {};
  const trackerSheetName = String(input.trackerSheetName || "").trim();
  const sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) throw new Error("Choose a tracker sheet.");

  const rawHeaders = getWalkthroughRawHeaders_(sheet);
  const normalized = rawHeaders.filter(Boolean).map(normalize_);
  if (normalized.some((value, index) => normalized.indexOf(value) !== index)) {
    throw new Error("Tracker headers must be unique before setup can continue.");
  }

  let recordIdHeader = String(input.recordIdHeader || "").trim();
  if (!recordIdHeader) recordIdHeader = "Name";
  let headers = getHeaders_(sheet);
  if (!headers[normalize_(recordIdHeader)]) {
    if (!input.createRecordIdColumn) {
      throw new Error('Record ID column "' + recordIdHeader + '" was not found.');
    }
    appendWalkthroughHeader_(sheet, recordIdHeader);
    headers = getHeaders_(sheet);
  }

  if (!headers.select) {
    if (!input.addSelectColumn) throw new Error('The required "Select" column is missing.');
    addSelectColumnToTracker_(sheet);
    headers = getHeaders_(sheet);
  }

  const templateSheet = ss.getSheetByName("Templates");
  const templates = templateSheet ? getWalkthroughValidTemplateKeys_(templateSheet) : [];
  if (!getRecipientFieldsForSendForm_(sheet, headers, templates).length) {
    if (!input.addEmailColumn) {
      throw new Error("Add or identify an email-recipient column before continuing.");
    }
    if (!headers.email) appendWalkthroughHeader_(sheet, "Email");
  }

  const result = saveSendMeBotSetup({
    trackerSheetName: trackerSheetName,
    recordIdHeader: recordIdHeader,
    addSelectColumn: true
  });
  updateSendMeBotInstallationRegistry_({ suggestedTrackerSheet: trackerSheetName });
  deleteUnusedTemplateTracker_(ss, trackerSheetName);
  return { result: result, context: getSendMeBotWalkthroughContext() };
}


function appendWalkthroughHeader_(sheet, header) {
  const value = String(header || "").trim();
  if (!value) throw new Error("Column name is required.");
  const headers = getHeaders_(sheet);
  if (headers[normalize_(value)]) return headers[normalize_(value)];
  const column = Math.max(sheet.getLastColumn(), 0) + 1;
  sheet.getRange(1, column).setValue(value);
  return column;
}


function deleteUnusedTemplateTracker_(ss, selectedTrackerName) {
  if (selectedTrackerName === "Tracker") return false;
  const templateTracker = ss.getSheetByName("Tracker");
  if (!templateTracker || !isTrackerEmptyBelowHeaders_(templateTracker)) return false;
  if (ss.getSheets().length <= 1) return false;
  ss.deleteSheet(templateTracker);
  return true;
}


function isTrackerEmptyBelowHeaders_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn < 1) return true;
  const headers = getHeaders_(sheet);
  const selectColumn = headers.select || 0;
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  return !values.some(row => row.some((value, index) =>
    index + 1 !== selectColumn && String(value || "").trim()
  ));
}


function saveWalkthroughSenderProfile(formData) {
  const result = saveSenderProfile_(formData || {}, { suppressToast: true });
  return { result: result, context: getSendMeBotWalkthroughContext() };
}


function saveWalkthroughTemplate(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Templates");
  const templates = sheet ? getWalkthroughValidTemplateKeys_(sheet) : [];
  const result = templates.length
    ? { title: "Template ready", message: templates[0] }
    : saveComposedTemplate(formData || {});
  return { result: result, context: getSendMeBotWalkthroughContext() };
}


function sendWalkthroughTestEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSendMeBotConfig_();
  const tracker = ss.getSheetByName(config.trackerSheetName);
  if (!tracker) throw new Error("Complete tracker setup before sending the test email.");
  const headers = getHeaders_(tracker);
  if (!headers.select || !headers[normalize_(config.recordIdHeader)]) {
    throw new Error("Complete tracker setup before sending the test email.");
  }

  const templateSheet = ss.getSheetByName("Templates");
  const templates = templateSheet ? getWalkthroughValidTemplateKeys_(templateSheet) : [];
  if (!templates.length) throw new Error("Create at least one email template before continuing.");
  if (!getRecipientFieldsForSendForm_(tracker, headers, templates).length) {
    throw new Error("Your tracker needs an email-recipient column before continuing.");
  }

  const authenticatedEmail = getAuthenticatedUserEmail_();
  const senderProfile = requireAuthenticatedSenderProfile_(ss, authenticatedEmail);
  const recordRow = getWalkthroughFirstRecordRow_(tracker, headers, config.recordIdHeader);
  const rowData = recordRow
    ? getRowData_(tracker, recordRow, headers)
    : buildWalkthroughSampleRowData_(headers, config.recordIdHeader, authenticatedEmail);
  const variableHeaders = getWalkthroughTestVariableHeaders_(headers, templates, config.recordIdHeader);
  const recordDisplayHeader = toDisplayHeader_(normalize_(config.recordIdHeader));
  const bodyRows = variableHeaders.map(header =>
    "<li><strong>" + escapeHtml_(header) + ":</strong> {{" + header + "}}</li>"
  ).join("");
  const controlledTemplate = {
    subject: "SendMeBot is ready — {{" + recordDisplayHeader + "}}",
    body: "<p>Hi {{Sender Name}},</p>" +
      "<p>Your tracker is connected. These values were merged using its live column names:</p>" +
      "<ul>" + bodyRows + "</ul>" +
      "<p>You can now select tracker rows and send with SendMeBot.</p>",
    attachmentLink: ""
  };
  const safeSenderProfile = {
    name: senderProfile.name,
    email: senderProfile.email,
    signatureText: "",
    signatureRichText: null
  };
  const payload = buildEmailPayload_(
    ss,
    tracker,
    recordRow || 2,
    headers,
    controlledTemplate,
    authenticatedEmail,
    { senderProfile: safeSenderProfile, rowData: rowData, imageAssets: {} }
  );

  const effectiveUser = getAuthenticatedUserEmail_();
  if (normalize_(effectiveUser) !== normalize_(authenticatedEmail)) {
    throw new Error("Sender does not match the authenticated Google account.");
  }
  MailApp.sendEmail({
    to: authenticatedEmail,
    subject: payload.subject,
    body: payload.plainBody,
    htmlBody: payload.htmlBody,
    name: payload.senderName
  });

  const now = new Date().toISOString();
  updateSendMeBotInstallationRegistry_({
    onboardingState: "complete",
    onboardingCompletedAt: now,
    testEmailSentAt: now
  });
  onOpen();
  return {
    recipient: authenticatedEmail,
    subject: payload.subject,
    usedTrackerRow: recordRow || 0,
    completedAt: now
  };
}


function getWalkthroughFirstRecordRow_(sheet, headers, recordIdHeader) {
  const recordColumn = recordIdHeader ? headers[normalize_(recordIdHeader)] : 0;
  if (!recordColumn || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, recordColumn, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = values.findIndex(row => String(row[0] || "").trim());
  return index === -1 ? 0 : index + 2;
}


function buildWalkthroughSampleRowData_(headers, recordIdHeader, authenticatedEmail) {
  const result = {};
  Object.keys(headers || {}).forEach(header => {
    const display = toDisplayHeader_(header);
    if (header === normalize_(recordIdHeader)) {
      result[display] = "SendMeBot test";
    } else if (header.indexOf("email") !== -1) {
      result[display] = authenticatedEmail;
    } else if (header === "select") {
      result[display] = "";
    } else {
      result[display] = "Sample " + display;
    }
  });
  return result;
}


function getWalkthroughTestVariableHeaders_(headers, templates, recordIdHeader) {
  const excluded = { select: true, status: true };
  (templates || []).forEach(template => {
    excluded[normalize_(template)] = true;
    excluded[normalize_(getTrackerHeaderForTemplate_(template))] = true;
  });
  const recordKey = normalize_(recordIdHeader);
  const ordered = [recordKey].concat(Object.keys(headers || {}).filter(header => header !== recordKey));
  return ordered
    .filter((header, index) =>
      header && (header === recordKey || !excluded[header]) && ordered.indexOf(header) === index
    )
    .slice(0, 5)
    .map(toDisplayHeader_);
}
