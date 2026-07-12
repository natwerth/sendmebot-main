// =============================================================================
// SendMeBot - Emails.gs
// Send form context, queueing, send/schedule processing, recipients, payloads
// =============================================================================

// =============================================================================
// Send form context / cache
// =============================================================================

function getSendFormOpenContext_(actionMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getTrackerSheet_();
  const templateSheet = ss.getSheetByName("Templates");

  if (!trackerSheet) throw new Error("Missing Tracker sheet.");
  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const headers = getHeaders_(trackerSheet);
  const selectedRows = getSelectedRows_(trackerSheet, headers);
  const templates = getTemplateKeys_(templateSheet);
  const templateTracking = {};
  templates.forEach(templateKey => {
    templateTracking[templateKey] = !!getTemplateStatusColumn_(headers, templateKey);
  });
  let authenticatedEmail = "";

  try {
    authenticatedEmail = getAuthenticatedUserEmail_();
  } catch (err) {
    return {
      actionMode: actionMode,
      selectedRowCount: selectedRows.length,
      authenticatedEmail: "",
      senderState: { status: "blocked", record: null, duplicateCount: 0 },
      triggerState: { status: "unknown", count: 0 },
      templates: templates,
      templateTracking: templateTracking,
      recipientFields: [],
      authorizedSenders: [],
      blockingError: err.message
    };
  }

  let senderState;
  let authorizedSenders;

  try {
    senderState = getSenderStateForEmail_(ss, authenticatedEmail);
    authorizedSenders = getAuthorizedSenderOptions_(ss);
  } catch (err) {
    return {
      actionMode: actionMode,
      selectedRowCount: selectedRows.length,
      authenticatedEmail: authenticatedEmail,
      senderState: { status: "blocked", record: null, duplicateCount: 0 },
      triggerState: { status: "unknown", count: 0 },
      templates: templates,
      templateTracking: templateTracking,
      recipientFields: [],
      authorizedSenders: [],
      blockingError: err.message || String(err)
    };
  }
  let triggerState;

  try {
    triggerState = getCurrentUserScheduledTriggerState_();
  } catch (err) {
    triggerState = {
      status: "unknown",
      count: 0,
      message: "Scheduled-trigger authorization will be checked when you schedule."
    };
  }

  return {
    actionMode: actionMode,
    selectedRowCount: selectedRows.length,
    authenticatedEmail: authenticatedEmail,
    senderState: toSenderBootstrapState_(senderState),
    triggerState: triggerState,
    templates: templates,
    templateTracking: templateTracking,
    recipientFields: getRecipientFieldsForSendForm_(trackerSheet, headers, templates),
    authorizedSenders: authorizedSenders,
    blockingError: ""
  };
}

function toSenderBootstrapState_(state, includeSignatureHtml) {
  const output = {
    status: state.status,
    email: state.email || "",
    duplicateCount: state.duplicateCount || 0,
    record: null
  };

  if (state.record) {
    output.record = {
      name: state.record.name || "",
      email: state.record.email || "",
      signature: state.record.signatureText || ""
    };
    if (includeSignatureHtml) {
      output.record.signatureHtml = getSenderSignatureEditorHtml_(state.record);
    }
  }

  return output;
}

function refreshFormCache() {
  const options = buildFormOptionsCache_();

  PropertiesService.getDocumentProperties().setProperty(
    "SENDMEBOT_FORM_OPTIONS",
    JSON.stringify(options)
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Form cache refreshed.",
    "SendMeBot",
    5
  );

  return options;
}

function getCachedFormOptions_() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("SENDMEBOT_FORM_OPTIONS");

  if (raw) {
    return JSON.parse(raw);
  }

  return refreshFormCache();
}

function buildFormOptionsCache_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const templates = getTemplateKeys_(templateSheet);

  return {
    templates: templates,
    refreshedAt: new Date().toISOString()
  };
}

// =============================================================================
// Recipient field detection for SendForm
// =============================================================================

function getRecipientFieldsForSendForm_(sheet, headers, templateKeys) {
  const templateKeyMap = {};

  (templateKeys || []).forEach(key => {
    templateKeyMap[normalize_(key)] = true;
  });

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const sampleSize = lastRow >= 2 ? Math.min(lastRow - 1, 100) : 0;

  // One read for sampled tracker rows.
  const values = sampleSize
    ? sheet.getRange(2, 1, sampleSize, lastCol).getDisplayValues()
    : [];

  // One read for actual display headers.
  const headerRow = lastCol
    ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    : [];

  const fields = [];

  Object.keys(headers).forEach(normalizedHeader => {
    const col = headers[normalizedHeader];
    const displayHeader = String(headerRow[col - 1] || "").trim() || toDisplayHeader_(normalizedHeader);
    const normalizedDisplayHeader = normalize_(displayHeader);

    // Never show core SendMeBot operational columns.
    if (
      normalizedDisplayHeader === "status" ||
      normalizedDisplayHeader === "select"
    ) {
      return;
    }

    // Exclude template-created status columns by exact template/header match.
    if (templateKeyMap[normalizedDisplayHeader]) return;

    // Exclude columns whose sampled values look like SendMeBot status tracking.
    if (isSendMeBotStatusColumnFromSample_(displayHeader, values, col)) return;

    const headerSuggestsEmail = normalizedDisplayHeader.indexOf("email") !== -1;

    let valuesSuggestEmail = false;

    for (let i = 0; i < values.length; i++) {
      const value = String(values[i][col - 1] || "").trim();

      if (looksLikeEmailList_(value)) {
        valuesSuggestEmail = true;
        break;
      }
    }

    // Main rule:
    // A recipient column must contain actual email-looking values.
    // This prevents "Hiring Manager", "Manager First Name", etc.
    if (!valuesSuggestEmail) return;

    // Secondary rule:
    // If the header does not say Email, still allow it only when the sampled
    // values clearly contain email addresses. This supports future CC/BCC helper
    // columns without making the detection stupid again.
    if (!headerSuggestsEmail && !valuesSuggestEmail) return;

    fields.push({
      header: displayHeader,
      label: getRecipientFieldLabel_(displayHeader)
    });
  });

  return fields.sort((a, b) => {
    if (normalize_(a.header) === "email") return -1;
    if (normalize_(b.header) === "email") return 1;
    return a.label.localeCompare(b.label);
  });
}


function getRecipientFieldLabel_(header) {
  let label = String(header || "")
    .replace(/\bemail address\b/ig, "")
    .replace(/\bemail\b/ig, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!label) return "Recipient";

  return label;
}


function isSendMeBotStatusColumnFromSample_(displayHeader, values, col) {
  const normalizedHeader = normalize_(displayHeader);

  if (
    normalizedHeader === "status" ||
    normalizedHeader === "select"
  ) {
    return true;
  }

  const columnValues = values
    .map(row => normalize_(row[col - 1]))
    .filter(Boolean);

  if (!columnValues.length) {
    return false;
  }

  const statusLikeCount = columnValues.filter(value =>
    value === "sent" ||
    value === "scheduled" ||
    value === "failed" ||
    value === "error: not sent" ||
    value === "sending..." ||
    value === "scheduling..." ||
    value === "sending scheduled email..." ||
    value.indexOf("sent on ") === 0 ||
    value.indexOf("scheduled for ") === 0
  ).length;

  return statusLikeCount > 0 && statusLikeCount >= Math.ceil(columnValues.length * 0.5);
}


function looksLikeEmailList_(value) {
  const raw = String(value || "").trim();

  if (!raw || raw.indexOf("@") === -1) return false;

  return raw
    .split(/[,;\n]/)
    .map(item => item.trim())
    .filter(Boolean)
    .some(item => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item));
}

// =============================================================================
// Recipient resolution for sending
// =============================================================================

function getRecipientConfigFromJob_(job) {
  return {
    sender: String(job.sender || "").trim(),
    toField: job.toField || { value: "Email", valueType: "field" },
    ccFields: Array.isArray(job.ccFields) ? job.ccFields : [],
    bccFields: Array.isArray(job.bccFields) ? job.bccFields : []
  };
}


function resolveRecipientsForRow_(sheet, row, headers, recipientConfig) {
  const config = recipientConfig || {
    sender: "",
    toField: { value: "Email", valueType: "field" },
    ccFields: [],
    bccFields: []
  };

  const to = getEmailValuesFromRecipientItem_(
    sheet,
    row,
    headers,
    config.toField,
    config.sender
  );

  const cc = getEmailValuesFromRecipientItems_(
    sheet,
    row,
    headers,
    config.ccFields || [],
    config.sender
  );

  const bcc = getEmailValuesFromRecipientItems_(
    sheet,
    row,
    headers,
    config.bccFields || [],
    config.sender
  );

  if (!to.length) {
    throw new Error("Missing recipient email.");
  }

  return dedupeRecipientGroups_(to, cc, bcc);
}


function getEmailValuesFromRecipientItems_(sheet, row, headers, items, selectedSender) {
  let emails = [];

  items.forEach(item => {
    emails = emails.concat(
      getEmailValuesFromRecipientItem_(sheet, row, headers, item, selectedSender)
    );
  });

  return emails;
}


function getEmailValuesFromRecipientItem_(sheet, row, headers, item, selectedSender) {
  if (!item) return [];

  if (typeof item === "string") {
    return getEmailValuesFromField_(sheet, row, headers, item);
  }

  const value = item.value ? String(item.value).trim() : "";
  const valueType = item.valueType ? String(item.valueType).trim() : "field";

  if (!value) return [];

  if (valueType === "email") {
    return parseEmailList_(value);
  }

  if (valueType === "selectedSender") {
    return parseEmailList_(selectedSender || "");
  }

  return getEmailValuesFromField_(sheet, row, headers, value);
}


function getEmailValuesFromField_(sheet, row, headers, fieldName) {
  const col = headers[normalize_(fieldName)];

  if (!col) return [];

  const value = sheet.getRange(row, col).getDisplayValue();

  return parseEmailList_(value);
}

function getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig) {
  const config = recipientConfig || {};
  const toField = config.toField || { value: "Email", valueType: "field" };

  if (toField.valueType === "selectedSender") {
    const senderProfile = getSenderProfile_(ss, config.sender || "");
    return senderProfile.name || config.sender || "";
  }

  if (toField.valueType === "email") {
    return String(toField.value || "").trim();
  }

  const toHeader = String(toField.value || "").trim();
  const normalizedToHeader = normalize_(toHeader);

  const nameFieldMap = {
    "student personal email": "Student Name",
    "student email": "Student Name",
    "email": "Student Name",

    "manager email": "Hiring Manager",
    "hiring manager email": "Hiring Manager",

    "mentor email": "Mentor"
  };

  const mappedNameField = nameFieldMap[normalizedToHeader];

  if (mappedNameField) {
    const mappedName = getCellValue_(sheet, row, headers, mappedNameField);
    if (mappedName) return mappedName;
  }

  const fallbackNameFields = [
    "Student Name",
    "Hiring Manager",
    "Mentor",
    "Name",
    "Full Name",
    "Candidate Name"
  ];

  for (let i = 0; i < fallbackNameFields.length; i++) {
    const value = getCellValue_(sheet, row, headers, fallbackNameFields[i]);
    if (value) return value;
  }

  return "";
}

function parseEmailList_(value) {
  return String(value || "")
    .split(/[,;\n]/)
    .map(item => item.trim())
    .filter(Boolean)
    .filter(item => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item));
}


function dedupeRecipientGroups_(to, cc, bcc) {
  const seen = {};

  const cleanTo = dedupeEmails_(to, seen);
  const cleanCc = dedupeEmails_(cc, seen);
  const cleanBcc = dedupeEmails_(bcc, seen);

  return {
    to: cleanTo.join(","),
    cc: cleanCc.join(","),
    bcc: cleanBcc.join(",")
  };
}


function dedupeEmails_(emails, seen) {
  const output = [];

  emails.forEach(email => {
    const key = normalize_(email);

    if (!key || seen[key]) return;

    seen[key] = true;
    output.push(email);
  });

  return output;
}



// =============================================================================
// Send form handoff
// =============================================================================

function requireAuthenticatedSenderProfile_(ss, submittedSender) {
  const authenticatedEmail = getAuthenticatedUserEmail_();
  const claimedEmail = String(submittedSender || "").trim();

  if (claimedEmail && normalize_(claimedEmail) !== normalize_(authenticatedEmail)) {
    throw new Error("Sender identity must match your authenticated Google account.");
  }

  const senderState = getSenderStateForEmail_(ss, authenticatedEmail);
  if (senderState.status === "duplicate") {
    throw new Error(
      "Duplicate sender profiles exist for " + authenticatedEmail +
      ". Remove the duplicate rows before sending."
    );
  }
  if (senderState.status !== "ready") {
    throw new Error("Set up your sender profile before sending.");
  }

  return senderState.record;
}

function prepareScheduledRows_(ss, sheet, rows, headers, recipientConfig, recordIdHeader) {
  const recordIdCol = headers[normalize_(recordIdHeader)];
  if (!recordIdCol) {
    throw new Error('Record ID column "' + recordIdHeader + '" was not found. Open SendMeBot → Setup.');
  }

  const lastRow = sheet.getLastRow();
  const allIds = lastRow > 1
    ? sheet.getRange(2, recordIdCol, lastRow - 1, 1).getDisplayValues().map(row => String(row[0] || "").trim())
    : [];
  const counts = {};
  allIds.forEach(value => {
    const key = normalize_(value);
    if (key) counts[key] = (counts[key] || 0) + 1;
  });
  const validRows = [];
  const failures = [];
  const preparedRows = {};

  rows.forEach(row => {
    try {
      const recordIdValue = String(sheet.getRange(row, recordIdCol).getDisplayValue() || "").trim();
      if (!recordIdValue) throw new Error('Record ID "' + recordIdHeader + '" is blank.');
      if (counts[normalize_(recordIdValue)] !== 1) {
        throw new Error(
          'Record ID "' + recordIdValue + '" is not unique in column "' + recordIdHeader + '".'
        );
      }
      const recipients = resolveRecipientsForRow_(sheet, row, headers, recipientConfig);
      preparedRows[row] = {
        recordIdHeader: recordIdHeader,
        recordIdValue: recordIdValue,
        recipients: recipients,
        name: recordIdValue
      };
      validRows.push(row);
    } catch (err) {
      failures.push({ row: row, error: err.message || String(err) });
    }
  });

  return { validRows: validRows, failures: failures, preparedRows: preparedRows };
}

function queueSendFormJob(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTrackerSheet_();

  if (!sheet) throw new Error("Missing Tracker sheet.");

  let headers = getHeaders_(sheet);
  const selectedRows = getSelectedRows_(sheet, headers);

  if (!selectedRows.length) {
    throw new Error("No selected rows found. Check at least one row in the Select column.");
  }

  const action = normalize_(formData.action);
  const templateKey = String(formData.template || "").trim();
  const toField = formData.toField || { value: "Email", valueType: "field" };
  const ccFields = Array.isArray(formData.ccFields) ? formData.ccFields : [];
  const bccFields = Array.isArray(formData.bccFields) ? formData.bccFields : [];
  const scheduledForIso = String(formData.scheduledForIso || "").trim();
  const scheduledTimeZone = String(formData.scheduledTimeZone || "").trim();
  const scheduledDisplayText = String(formData.scheduledDisplayText || "").trim();
  const createTrackingColumn = formData.createTrackingColumn === true;

  if (!templateKey) throw new Error("Template is required.");
  if (!toField) throw new Error("Recipient field is required.");

  if (action !== "send_now" && action !== "schedule") {
    throw new Error("Unknown action: " + formData.action);
  }

  let trackTemplateStatus = !!getTemplateStatusColumn_(headers, templateKey);
  if (!trackTemplateStatus && createTrackingColumn) {
    ensureTrackerColumnForTemplate_(templateKey);
    headers = getHeaders_(sheet);
    trackTemplateStatus = !!getTemplateStatusColumn_(headers, templateKey);
    if (!trackTemplateStatus) throw new Error("Tracking column creation could not be verified.");
  }

  if (action === "schedule" && !trackTemplateStatus) {
    throw new Error("Create a tracking column before scheduling this template.");
  }

  if (action === "schedule") {
    markSchedulingRowsStarted_(sheet, headers, selectedRows, templateKey);
  }

  let senderProfile;
  try {
    if (action === "schedule") {
      validateScheduledInstant_(scheduledForIso, new Date());
      if (!scheduledTimeZone) throw new Error("Browser timezone is required for scheduling.");
      if (!scheduledDisplayText) throw new Error("Scheduled display time is required.");
    }
    senderProfile = requireAuthenticatedSenderProfile_(ss, formData.sender);
  } catch (requestErr) {
    if (action === "schedule") {
      markSchedulingRequestError_(sheet, headers, selectedRows, templateKey, requestErr.message || requestErr);
    }
    throw requestErr;
  }
  const sender = senderProfile.email;
  const recipientConfig = {
    sender: sender,
    toField: toField,
    ccFields: ccFields,
    bccFields: bccFields
  };
  let rowsToProcess = selectedRows.slice();
  let preflightFailures = [];
  let preparedRows = {};

  if (action === "schedule") {
    const config = getSendMeBotConfig_();
    let preflight;
    try {
      preflight = prepareScheduledRows_(
        ss,
        sheet,
        selectedRows,
        headers,
        recipientConfig,
        config.recordIdHeader
      );
    } catch (preflightErr) {
      markSchedulingRequestError_(
        sheet, headers, selectedRows, templateKey, preflightErr.message || preflightErr
      );
      throw preflightErr;
    }
    rowsToProcess = preflight.validRows;
    preflightFailures = preflight.failures;
    preparedRows = preflight.preparedRows;

    if (!rowsToProcess.length) {
      preflightFailures.forEach(item => {
        markScheduledRowError_(sheet, headers, item.row, templateKey, item.error);
        clearSelectedRow_(sheet, headers, item.row);
      });
      SpreadsheetApp.flush();
      return {
        title: "Scheduling failed",
        action: action,
        attempted: selectedRows.length,
        successful: 0,
        sent: 0,
        scheduled: 0,
        failed: preflightFailures.length,
        successfulRows: [],
        failedRows: preflightFailures.map(item => item.row),
        errors: preflightFailures,
        locked: false,
        success: false,
        message: "No emails were scheduled. Fix the selected rows and try again."
      };
    }

    try {
      ensureCurrentUserScheduledTrigger_();
    } catch (triggerErr) {
      markSchedulingRequestError_(
        sheet, headers, selectedRows, templateKey, triggerErr.message || triggerErr
      );
      throw triggerErr;
    }
  }

  const job = {
    jobId: makeJobId_(),
    createdAt: new Date().toISOString(),
    createdBy: sender,
    action: action,
    template: templateKey,
    sender: sender,
    toField: toField,
    ccFields: ccFields,
    bccFields: bccFields,
    scheduledForIso: scheduledForIso,
    scheduledTimeZone: scheduledTimeZone,
    scheduledDisplayText: scheduledDisplayText,
    rows: rowsToProcess,
    attemptedRows: selectedRows,
    preflightFailures: preflightFailures,
    preparedScheduledRows: preparedRows,
    trackTemplateStatus: trackTemplateStatus
  };

  let result;
  if (action === "schedule") {
    try {
      result = processSchedulingJobNow_(job);
    } catch (processingErr) {
      markSchedulingRequestError_(
        sheet, headers, selectedRows, templateKey, processingErr.message || processingErr
      );
      throw processingErr;
    }
  } else {
    saveQueuedJob_(job);
    // Current modeless flow: process immediately after handoff.
    // Backup menu/trigger can still process orphaned immediate-send jobs.
    result = processQueuedJobs_();
  }
  result = result || {
    attempted: selectedRows.length,
    successful: 0,
    sent: 0,
    scheduled: 0,
    failed: 0,
    successfulRows: [],
    failedRows: [],
    errors: [],
    skipped: 0,
    locked: false
  };

  const successCount = (result.sent || 0) + (result.scheduled || 0);
  const failedCount = result.failed || 0;

  if (action === "schedule" && result.locked) {
    markSchedulingRequestError_(
      sheet, headers, selectedRows, templateKey, "Scheduling is busy. Please try again."
    );
  }

  return {
    title: failedCount ? "Request completed with errors" : "Request completed",
    action: action,
    attempted: result.attempted || selectedRows.length,
    successful: successCount,
    sent: result.sent || 0,
    scheduled: result.scheduled || 0,
    failed: failedCount,
    successfulRows: result.successfulRows || [],
    failedRows: result.failedRows || [],
    errors: result.errors || [],
    skipped: result.skipped || 0,
    locked: result.locked || false,
    success: failedCount === 0 && successCount > 0,
    message:
      "Selected: " + selectedRows.length +
      ". Sent: " + (result.sent || 0) +
      ". Scheduled: " + (result.scheduled || 0) +
      ". Failed: " + failedCount +
      "."
  };
}

// =============================================================================
// Job queue
// =============================================================================

function makeJobId_() {
  return "job_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000000);
}


function getQueuedJobs_() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("SENDMEBOT_QUEUED_JOBS");

  if (!raw) return [];

  try {
    return JSON.parse(raw) || [];
  } catch (err) {
    return [];
  }
}


function saveQueuedJobs_(jobs) {
  PropertiesService.getDocumentProperties().setProperty(
    "SENDMEBOT_QUEUED_JOBS",
    JSON.stringify(jobs)
  );
}


function saveQueuedJob_(job) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const jobs = getQueuedJobs_();
    jobs.push(job);
    saveQueuedJobs_(jobs);
  } finally {
    lock.releaseLock();
  }
}


function processQueuedJobs() {
  return processQueuedJobs_();
}

function processSchedulingJobNow_(job) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    return {
      attempted: (job.attemptedRows || []).length,
      successful: 0,
      sent: 0,
      scheduled: 0,
      failed: 0,
      successfulRows: [],
      failedRows: [],
      errors: [],
      skipped: 0,
      locked: true
    };
  }
  try {
    return processOneQueuedJob_(job);
  } finally {
    lock.releaseLock();
  }
}

function processQueuedJobs_() {
  const lock = LockService.getDocumentLock();

  const summary = {
    attempted: 0,
    successful: 0,
    sent: 0,
    scheduled: 0,
    failed: 0,
    successfulRows: [],
    failedRows: [],
    errors: [],
    skipped: 0,
    locked: false
  };

  if (!lock.tryLock(30000)) {
    summary.locked = true;
    return summary;
  }

  try {
    const jobs = getQueuedJobs_();

    if (!jobs.length) {
      return summary;
    }

    saveQueuedJobs_([]);

    const failedToProcess = [];

    jobs.forEach(job => {
      try {
        const result = processOneQueuedJob_(job) || {};

        summary.attempted += result.attempted || 0;
        summary.successful += result.successful || 0;
        summary.sent += result.sent || 0;
        summary.scheduled += result.scheduled || 0;
        summary.failed += result.failed || 0;
        summary.successfulRows = summary.successfulRows.concat(result.successfulRows || []);
        summary.failedRows = summary.failedRows.concat(result.failedRows || []);
        summary.errors = summary.errors.concat(result.errors || []);
        summary.skipped += result.skipped || 0;
      } catch (err) {
        Logger.log("JOB ERROR " + job.jobId + ": " + err.message);
        failedToProcess.push(job);
        summary.failed++;
      }
    });

    if (failedToProcess.length) {
      const remainingJobs = getQueuedJobs_().concat(failedToProcess);
      saveQueuedJobs_(remainingJobs);
    }

    return summary;

  } finally {
    lock.releaseLock();
  }
}


function processOneQueuedJob_(job) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTrackerSheet_();

  if (!sheet) throw new Error("Missing Tracker sheet.");

  let headers = getHeaders_(sheet);
  const rows = job.rows || [];
  const recipientConfig = getRecipientConfigFromJob_(job);

  let sentCount = 0;
  let scheduledCount = 0;
  const preflightFailures = Array.isArray(job.preflightFailures)
    ? job.preflightFailures.slice()
    : [];
  let failedCount = preflightFailures.length;
  let skippedCount = 0;
  const successfulRows = [];
  const failedRows = preflightFailures.map(item => item.row);
  const errors = preflightFailures.slice();
  const attemptedRows = Array.isArray(job.attemptedRows) ? job.attemptedRows : rows;

  if (job.action === "send_now") {
    const templateSheet = ss.getSheetByName("Templates");
    if (!templateSheet) throw new Error("Missing Templates sheet.");
    const hasTrackingFlag = Object.prototype.hasOwnProperty.call(job, "trackTemplateStatus");
    const trackingOptions = {
      enabled: hasTrackingFlag ? job.trackTemplateStatus === true : true,
      allowLegacyCreate: !hasTrackingFlag
    };
    rows.forEach(row => {
      let result;
      try {
        result = sendOneRowNow_(
          ss,
          sheet,
          row,
          headers,
          job.template,
          job.sender,
          recipientConfig,
          trackingOptions
        );
      } catch (err) {
        result = { status: "failed", row: row, error: err.message || String(err) };
      }

      if (result.status === "sent") {
        sentCount++;
        successfulRows.push(row);
      } else {
        failedCount++;
        failedRows.push(row);
        errors.push({ row: row, error: result.error || "Email failed." });
      }
      try {
        clearSelectedRow_(sheet, headers, row);
      } catch (checkboxErr) {
        Logger.log("CHECKBOX CLEAR ERROR row " + row + ": " + checkboxErr.message);
      }
    });
    try {
      SpreadsheetApp.flush();
    } catch (flushErr) {
      Logger.log("SEND CHECKBOX FLUSH ERROR: " + flushErr.message);
    }

    ss.toast(
      "Send complete. Sent: " + sentCount + ". Failed: " + failedCount + ".",
      "SendMeBot",
      8
    );

    return {
      attempted: attemptedRows.length,
      successful: sentCount,
      sent: sentCount,
      scheduled: 0,
      failed: failedCount,
      successfulRows: successfulRows,
      failedRows: failedRows,
      errors: errors,
      skipped: skippedCount
    };
  }

  if (job.action === "schedule") {
    const scheduledDate = parseScheduledInstant_(job.scheduledForIso);
    const preparedRows = job.preparedScheduledRows || {};

    preflightFailures.forEach(item => {
      try {
        markScheduledRowError_(sheet, headers, item.row, job.template, item.error);
        clearSelectedRow_(sheet, headers, item.row);
      } catch (checkboxErr) {
        Logger.log("CHECKBOX CLEAR ERROR row " + item.row + ": " + checkboxErr.message);
      }
    });

    rows.forEach(row => {
      let result;
      try {
        result = scheduleOneRow_(
          ss,
          sheet,
          row,
          headers,
          job.template,
          job.sender,
          scheduledDate,
          job.scheduledTimeZone,
          job.scheduledDisplayText,
          recipientConfig,
          preparedRows[row]
        );
      } catch (err) {
        result = { status: "failed", row: row, error: err.message || String(err) };
      }

      if (result.status === "scheduled") {
        scheduledCount++;
        successfulRows.push(row);
      } else {
        failedCount++;
        failedRows.push(row);
        errors.push({ row: row, error: result.error || "Scheduling failed." });
      }
      try {
        clearSelectedRow_(sheet, headers, row);
      } catch (checkboxErr) {
        Logger.log("CHECKBOX CLEAR ERROR row " + row + ": " + checkboxErr.message);
      }
    });
    try {
      SpreadsheetApp.flush();
    } catch (flushErr) {
      Logger.log("SCHEDULE CHECKBOX FLUSH ERROR: " + flushErr.message);
    }

    ss.toast(
      "Scheduling complete. Scheduled: " + scheduledCount + ". Failed: " + failedCount + ".",
      "SendMeBot",
      8
    );

    return {
      attempted: attemptedRows.length,
      successful: scheduledCount,
      sent: 0,
      scheduled: scheduledCount,
      failed: failedCount,
      successfulRows: successfulRows,
      failedRows: failedRows,
      errors: errors,
      skipped: skippedCount
    };
  }

  throw new Error("Unknown queued action: " + job.action);
}

function markSchedulingRowsStarted_(sheet, headers, rows, templateKey) {
  if (!rows.length) return;

  const statusHeaderCol = getTemplateStatusColumn_(headers, templateKey);
  if (!statusHeaderCol) throw new Error("Template tracking column is missing.");
  const templateRanges = rows.map(row => sheet.getRange(row, statusHeaderCol).getA1Notation());
  sheet.getRangeList(templateRanges).setValue("Scheduling...");

  SpreadsheetApp.flush();
}

function shortenTrackerError_(message) {
  const text = String(message || "Scheduling failed.").replace(/^Error:\s*/i, "").trim();
  return text.length > 120 ? text.slice(0, 117) + "..." : text;
}

function markScheduledRowError_(sheet, headers, row, templateKey, message) {
  const templateCol = getTemplateStatusColumn_(headers, templateKey);
  if (!templateCol) return;
  sheet.getRange(row, templateCol).setValue("Error: " + shortenTrackerError_(message));
}

function markSchedulingRequestError_(sheet, headers, rows, templateKey, message) {
  rows.forEach(row => markScheduledRowError_(sheet, headers, row, templateKey, message));
  SpreadsheetApp.flush();
}

function getCurrentUserScheduledTriggerState_(runtime) {
  const options = runtime || {};
  const scriptApp = options.scriptApp || ScriptApp;
  const triggers = scriptApp.getProjectTriggers();
  const matches = triggers.filter(trigger =>
    trigger.getHandlerFunction() === "sendScheduledEmails"
  );

  if (matches.length > 1) {
    return {
      status: "duplicate",
      count: matches.length,
      message: "Multiple scheduled-email triggers are installed for your account."
    };
  }

  if (!matches.length) {
    return {
      status: "missing",
      count: 0,
      message: "A five-minute scheduled-email trigger will be created when you schedule."
    };
  }

  // Apps Script exposes handler and event type, but not a trigger's interval.
  // One clock trigger with this exact handler is the strongest runtime check;
  // triggers created here always use the approved five-minute cadence.
  const clockEvent = scriptApp.EventType && scriptApp.EventType.CLOCK;
  if (clockEvent && matches[0].getEventType() !== clockEvent) {
    return {
      status: "invalid",
      count: 1,
      message: "The existing sendScheduledEmails trigger is not time-driven."
    };
  }

  return {
    status: "ready",
    count: 1,
    message: "Scheduled-email processing is ready for your account."
  };
}

function ensureCurrentUserScheduledTrigger_(runtime) {
  const options = runtime || {};
  const scriptApp = options.scriptApp || ScriptApp;
  const lock = options.lock || LockService.getUserLock();
  lock.waitLock(30000);

  try {
    const initialState = getCurrentUserScheduledTriggerState_({ scriptApp: scriptApp });
    if (initialState.status === "ready") return initialState;

    if (initialState.status === "duplicate") {
      throw new Error(
        "Multiple sendScheduledEmails triggers exist for your account. " +
        "Open Apps Script → Triggers and remove the duplicate before scheduling."
      );
    }

    if (initialState.status === "invalid") {
      throw new Error(
        "Your sendScheduledEmails trigger is not a valid time-driven trigger. " +
        "Open Apps Script → Triggers and correct it before scheduling."
      );
    }

    try {
      scriptApp.newTrigger("sendScheduledEmails")
        .timeBased()
        .everyMinutes(5)
        .create();
    } catch (err) {
      throw new Error(
        "Scheduled sending could not be authorized or enabled. " +
        "Authorize SendMeBot and try again. If prompted outside the dialog, " +
        "open Apps Script and run setupScheduledEmailTrigger once. Details: " +
        (err.message || err)
      );
    }

    const verifiedState = getCurrentUserScheduledTriggerState_({ scriptApp: scriptApp });
    if (verifiedState.status !== "ready") {
      throw new Error(
        "Scheduled-trigger creation could not be verified. No emails were scheduled."
      );
    }

    return verifiedState;
  } finally {
    lock.releaseLock();
  }
}

function setupScheduledEmailTrigger() {
  const state = ensureCurrentUserScheduledTrigger_();
  SpreadsheetApp.getActiveSpreadsheet().toast(state.message, "SendMeBot", 5);
  return state;
}

function setupJobProcessorTrigger() {
  const functionName = "processQueuedJobs";

  const alreadyExists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === functionName
  );

  if (alreadyExists) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Job processor trigger is already set up.",
      "SendMeBot",
      5
    );
    return;
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Job processor is set up to check queued jobs every minute.",
    "SendMeBot",
    5
  );
}

// =============================================================================
// Immediate send
// =============================================================================

function sendOneRowNow_(ss, sheet, row, headers, templateKey, sender, recipientConfig, trackingOptions) {
  const templateSheet = ss.getSheetByName("Templates");
  const tracking = trackingOptions || { enabled: true, allowLegacyCreate: true };

  let recipients = {
    to: "",
    cc: "",
    bcc: ""
  };

  let email = "";

  const name = getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig);

  try {
    let templateStatusCol = null;
    if (tracking.enabled) {
      templateStatusCol = getTemplateStatusColumn_(getHeaders_(sheet), templateKey);
      if (!templateStatusCol && tracking.allowLegacyCreate) {
        templateStatusCol = ensureTrackerColumnForTemplate_(templateKey);
      }
      if (!templateStatusCol) throw new Error("Template tracking column is missing.");
      sheet.getRange(row, templateStatusCol).setValue("Sending...");
    }

    recipients = resolveRecipientsForRow_(
      sheet,
      row,
      headers,
      recipientConfig
    );

    email = recipients.to;

    if (!email) throw new Error("Missing recipient email.");

    const template = getTemplateByKey_(templateSheet, templateKey);
    const payload = buildEmailPayload_(ss, sheet, row, headers, template, sender);

    // Revalidate the execution authority immediately before MailApp.
    const effectiveUser = getAuthenticatedUserEmail_();
    if (normalize_(effectiveUser) !== normalize_(sender)) {
      throw new Error("Sender does not match the authenticated Google account.");
    }

    MailApp.sendEmail({
      to: recipients.to,
      cc: recipients.cc,
      bcc: recipients.bcc,
      subject: payload.subject,
      body: payload.plainBody,
      htmlBody: payload.htmlBody,
      inlineImages: payload.inlineImages,
      attachments: payload.attachments,
      name: payload.senderName
    });

    stampTemplateColumn_(sheet, row, getHeaders_(sheet), templateKey, tracking.enabled);

    logSentEmail_(ss, {
      name,
      email,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: templateKey,
      sender,
      subject: payload.subject,
      status: "Sent",
      message: "Email sent successfully.",
      sourceRow: row,
      body: payload.plainBody,
      attachments: payload.attachmentNames || ""
    });

    return { status: "sent", row: row, error: "" };

  } catch (err) {
    try {
      stampTemplateFailure_(sheet, row, getHeaders_(sheet), templateKey, tracking.enabled);

      logSentEmail_(ss, {
        name,
        email,
        cc: recipients.cc,
        bcc: recipients.bcc,
        template: templateKey,
        sender,
        subject: "",
        status: "Failed",
        message: err.message || "Email failed.",
        sourceRow: row,
        body: "",
        attachments: ""
      });
    } catch (failureLogErr) {
      Logger.log("SEND FAILURE LOG ERROR row " + row + ": " + failureLogErr.message);
    }

    Logger.log("SEND ERROR row " + row + ": " + err.message);
    return { status: "failed", row: row, error: err.message || "Email failed." };
  }
}

// =============================================================================
// Schedule selected / Sent-sheet queue
// =============================================================================

function parseScheduledInstant_(isoValue) {
  const raw = String(isoValue || "").trim();
  const scheduledDate = new Date(raw);

  if (!raw || isNaN(scheduledDate.getTime())) {
    throw new Error("Invalid scheduled datetime.");
  }

  return scheduledDate;
}

function validateScheduledInstant_(isoValue, nowValue) {
  const scheduledDate = parseScheduledInstant_(isoValue);
  const now = nowValue instanceof Date ? nowValue : new Date(nowValue || new Date());

  if (isNaN(now.getTime())) throw new Error("Invalid current datetime.");
  if (scheduledDate.getTime() <= now.getTime()) {
    throw new Error("Scheduled time must be in the future.");
  }

  return scheduledDate;
}

function getLiveDeliveryMetadata_(sheet, row, statusColumn, prepared, audit) {
  return {
    version: 3,
    sourceSheet: sheet.getName(),
    sourceRow: row,
    sourceStatusColumn: statusColumn,
    recordIdHeader: prepared.recordIdHeader,
    recordIdValue: prepared.recordIdValue,
    scheduledTimeZone: audit.scheduledTimeZone,
    scheduledDisplayText: audit.scheduledDisplayText
  };
}

function scheduleOneRow_(
  ss,
  sheet,
  row,
  headers,
  templateKey,
  sender,
  scheduledDate,
  scheduledTimeZone,
  scheduledDisplayText,
  recipientConfig,
  preparedRow
) {
  let queueRowWritten = false;
  let queueRow = 0;

  try {
    if (!(scheduledDate instanceof Date) || isNaN(scheduledDate.getTime())) {
      throw new Error("Invalid scheduled datetime.");
    }

    if (normalize_(getAuthenticatedUserEmail_()) !== normalize_(sender)) {
      throw new Error("Sender does not match the authenticated Google account.");
    }

    const prepared = preparedRow;
    if (!prepared || !prepared.recordIdHeader || !prepared.recordIdValue) {
      throw new Error("Scheduled row is missing its prepared source-record identity.");
    }
    const statusHeaderCol = getTemplateStatusColumn_(headers, templateKey);
    if (!statusHeaderCol) throw new Error("Template tracker column could not be verified.");
    const recipients = prepared.recipients;
    if (!recipients || !recipients.to) throw new Error("Scheduled row is missing Recipient.");
    const metadata = getLiveDeliveryMetadata_(sheet, row, statusHeaderCol, prepared, {
      scheduledTimeZone: scheduledTimeZone,
      scheduledDisplayText: scheduledDisplayText
    });

    queueRow = logSentEmail_(ss, {
      timestamp: new Date(),
      status: "Scheduled",
      scheduledFor: new Date(scheduledDate.getTime()),
      processedAt: "",
      message: "Email scheduled for " + scheduledDisplayText + ".",
      name: prepared.name,
      email: recipients.to,
      sender: sender,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: templateKey,
      subject: "",
      body: "",
      attachments: "",
      logNote: JSON.stringify(metadata)
    });
    queueRowWritten = true;

    const visibleScheduledDate = Utilities.formatDate(
      scheduledDate,
      ss.getSpreadsheetTimeZone(),
      "M/d"
    );
    sheet.getRange(row, statusHeaderCol).setValue("Scheduled for " + visibleScheduledDate);
    return { status: "scheduled", row: row, error: "" };
  } catch (err) {
    if (queueRowWritten) {
      try {
        const sentSheet = ss.getSheetByName("Sent");
        const sentHeaders = getRequiredSentQueueHeaders_(sentSheet);
        const queueFailureMessage =
          "Scheduling failed after the queue row was written: " + (err.message || err);
        sentSheet.getRange(queueRow, sentHeaders["status"]).setValue("Failed");
        sentSheet.getRange(queueRow, sentHeaders["processed at"]).setValue(new Date());
        sentSheet.getRange(queueRow, sentHeaders["message"]).setValue(queueFailureMessage);
        sentSheet.getRange(queueRow, sentHeaders["log note"]).setValue(queueFailureMessage);
      } catch (queueErr) {
        Logger.log("SCHEDULE QUEUE FINALIZATION ERROR row " + row + ": " + queueErr.message);
      }
    }

    try {
      markScheduledRowError_(sheet, getHeaders_(sheet), row, templateKey, err.message || err);
    } catch (statusErr) {
      Logger.log("SCHEDULE STATUS ERROR row " + row + ": " + statusErr.message);
    }

    Logger.log("SCHEDULE ERROR row " + row + ": " + (err.message || err));
    return { status: "failed", row: row, error: err.message || "Scheduling failed." };
  }
}

// The trigger event is intentionally ignored. There is no force-send path.
function sendScheduledEmails() {
  return processScheduledEmails_();
}

function getRequiredSentQueueHeaders_(sheet) {
  const headers = getHeaders_(sheet);
  [
    "status", "scheduled for", "processed at", "message", "recipient",
    "sender", "cc", "bcc", "template", "subject", "email body", "attachments", "log note"
  ].forEach(header => {
    if (!headers[header]) throw new Error("Missing Sent header: " + header + ".");
  });
  return headers;
}

function parseScheduledLogNote_(value) {
  let metadata;

  try {
    metadata = JSON.parse(String(value || ""));
  } catch (err) {
    throw new Error("Scheduled row has invalid delivery metadata.");
  }

  if (!metadata || metadata.version !== 3) {
    const version = metadata && metadata.version !== undefined ? metadata.version : "unknown";
    throw new Error(
      "Unsupported legacy scheduled payload (version " + version + "). Reschedule this email."
    );
  }

  return metadata;
}

function updateSentQueueRow_(sheet, row, rowValues, headers, updates) {
  Object.keys(updates).forEach(header => {
    const col = headers[normalize_(header)];
    if (col) rowValues[col - 1] = updates[header];
  });
  sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
}

function resolveScheduledSourceRecord_(ss, metadata, templateKey, cache) {
  const sourceSheet = ss.getSheetByName(String(metadata.sourceSheet || ""));
  if (!sourceSheet) {
    return { state: "orphaned", reason: 'source sheet "' + metadata.sourceSheet + '" no longer exists' };
  }
  const recordIdHeader = String(metadata.recordIdHeader || "").trim();
  const recordIdValue = String(metadata.recordIdValue || "").trim();
  if (!recordIdHeader || !recordIdValue) {
    return { state: "orphaned", reason: "source record identity is missing" };
  }
  const currentHeaders = getHeaders_(sourceSheet);
  const idCol = currentHeaders[normalize_(recordIdHeader)];
  if (!idCol) {
    return {
      state: "orphaned",
      reason: 'record ID column "' + recordIdHeader + '" no longer exists in "' + sourceSheet.getName() + '"'
    };
  }
  const cacheKey = sourceSheet.getName() + "\n" + normalize_(recordIdHeader);
  const lookupCache = cache || {};
  if (!lookupCache[cacheKey]) {
    const lastRow = sourceSheet.getLastRow();
    const values = lastRow > 1
      ? sourceSheet.getRange(2, idCol, lastRow - 1, 1).getDisplayValues()
      : [];
    const rowsById = {};
    values.forEach((value, index) => {
      const key = normalize_(value[0]);
      if (!key) return;
      if (!rowsById[key]) rowsById[key] = [];
      rowsById[key].push(index + 2);
    });
    lookupCache[cacheKey] = rowsById;
  }
  const matches = lookupCache[cacheKey][normalize_(recordIdValue)] || [];
  if (!matches.length) {
    return { state: "orphaned", reason: 'source record "' + recordIdValue + '" could not be found' };
  }
  if (matches.length > 1) {
    return {
      state: "failed",
      reason: 'record ID "' + recordIdValue + '" matched multiple Tracker rows; sending was blocked'
    };
  }
  const sourceRow = matches[0];
  const sourceCol = getTemplateStatusColumn_(currentHeaders, templateKey);
  if (!sourceCol) {
    return { state: "orphaned", reason: "template-specific source status column no longer exists" };
  }
  const currentValue = sourceSheet.getRange(sourceRow, sourceCol).getDisplayValue();
  if (currentValue.indexOf("Scheduled for ") !== 0) {
    return { state: "cancelled", reason: "tracker status was cleared or changed" };
  }
  return {
    state: "valid",
    sheet: sourceSheet,
    row: sourceRow,
    headers: currentHeaders,
    statusCol: sourceCol
  };
}

function updateScheduledSourceTracker_(resolution, templateKey, outcome, message) {
  if (!resolution || resolution.state !== "valid") return false;
  const sourceSheet = resolution.sheet;
  const sourceRow = resolution.row;
  const trackerHeaders = resolution.headers;
  if (outcome === "Sent") {
    stampTemplateColumn_(sourceSheet, sourceRow, trackerHeaders, templateKey);
  } else {
    markScheduledRowError_(sourceSheet, trackerHeaders, sourceRow, templateKey, message);
  }
  return true;
}

function getTemplateTokens_(value) {
  const tokens = [];
  const pattern = /\{\{(.*?)\}\}/g;
  let match;
  while ((match = pattern.exec(String(value || ""))) !== null) {
    const token = String(match[1] || "").trim();
    if (token && tokens.indexOf(token) === -1) tokens.push(token);
  }
  return tokens;
}

function validateLiveRenderingTokens_(template, senderProfile, rowData, imageAssets) {
  const allowedText = Object.assign({}, rowData, {
    "Sender Name": senderProfile.name || "",
    "Sender Email": senderProfile.email || ""
  });
  const imageKeys = {};
  Object.keys(imageAssets || {}).forEach(key => { imageKeys[normalizeTemplateKey_(key)] = true; });
  const validate = function(values, allowSignature) {
    values.forEach(value => getTemplateTokens_(value).forEach(token => {
      if (allowSignature && token === "Sender Signature") return;
      if (Object.prototype.hasOwnProperty.call(allowedText, token)) return;
      if (imageKeys[normalizeTemplateKey_(token)]) return;
      throw new Error(
        'Template variable "{{' + token + '}}" could not be resolved because no current Tracker column or image has that name.'
      );
    }));
  };
  validate([senderProfile.signatureText || ""], false);
  validate([template.subject || "", template.body || ""], true);
}

function processScheduledEmails_(options) {
  const runtime = options || {};
  const lock = runtime.lock || LockService.getScriptLock();
  const summary = {
    sent: 0,
    failed: 0,
    cancelled: 0,
    orphaned: 0,
    skipped: 0,
    locked: false
  };

  if (!lock.tryLock(30000)) {
    summary.locked = true;
    return summary;
  }

  try {
    const ss = runtime.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const sentSheet = runtime.sentSheet || ss.getSheetByName("Sent");
    if (!sentSheet) throw new Error("Missing Sent sheet.");

    const headers = getRequiredSentQueueHeaders_(sentSheet);
    const values = sentSheet.getDataRange().getValues();
    const now = runtime.now ? new Date(runtime.now) : new Date();
    const getEffectiveUserEmail = runtime.getEffectiveUserEmail || function() {
      return Object.prototype.hasOwnProperty.call(runtime, "effectiveUser")
        ? String(runtime.effectiveUser || "")
        : getAuthenticatedUserEmail_();
    };
    const sendEmail = runtime.sendEmail || function(message) { MailApp.sendEmail(message); };
    const sourceLookupCache = {};
    let effectiveUser = "";

    try {
      effectiveUser = String(getEffectiveUserEmail() || "").trim();
    } catch (identityErr) {
      Logger.log("SCHEDULED SEND SKIPPED: " + (identityErr.message || identityErr));
      summary.skipped = Math.max(values.length - 1, 0);
      return summary;
    }

    if (!effectiveUser) {
      Logger.log("SCHEDULED SEND SKIPPED: effective trigger owner could not be determined.");
      summary.skipped = Math.max(values.length - 1, 0);
      return summary;
    }

    for (let r = 1; r < values.length; r++) {
      const rowValues = values[r].slice();
      const row = r + 1;
      if (rowValues[headers["status"] - 1] !== "Scheduled") continue;

      const sender = String(rowValues[headers["sender"] - 1] || "").trim();
      if (!sender || normalize_(sender) !== normalize_(effectiveUser)) {
        summary.skipped++;
        continue;
      }

      const scheduledFor = rowValues[headers["scheduled for"] - 1];

      if (!(scheduledFor instanceof Date) || isNaN(scheduledFor.getTime())) {
        const invalidMessage = "Scheduled For must be a valid spreadsheet datetime.";
        updateSentQueueRow_(sentSheet, row, rowValues, headers, {
          "Status": "Failed",
          "Processed At": new Date(now.getTime()),
          "Message": invalidMessage,
          "Log Note": invalidMessage
        });
        summary.failed++;
        continue;
      }

      if (scheduledFor.getTime() > now.getTime()) {
        summary.skipped++;
        continue;
      }

      let delivered = false;
      let metadata = null;
      let sourceResolution = null;
      let renderedPayload = null;

      try {
        metadata = parseScheduledLogNote_(rowValues[headers["log note"] - 1]);
        sourceResolution = resolveScheduledSourceRecord_(
          ss,
          metadata,
          rowValues[headers["template"] - 1],
          sourceLookupCache
        );

        if (sourceResolution.state === "orphaned") {
          const orphanedMessage =
            "Scheduled email was not sent because " + sourceResolution.reason + ".";
          updateSentQueueRow_(sentSheet, row, rowValues, headers, {
            "Status": "Orphaned",
            "Processed At": new Date(now.getTime()),
            "Message": orphanedMessage,
            "Log Note": orphanedMessage
          });
          summary.orphaned++;
          continue;
        }

        if (sourceResolution.state === "cancelled") {
          const cancelledMessage =
            "Scheduled email was cancelled because the tracker status was cleared or changed.";
          updateSentQueueRow_(sentSheet, row, rowValues, headers, {
            "Status": "Cancelled",
            "Processed At": new Date(now.getTime()),
            "Message": cancelledMessage,
            "Log Note": cancelledMessage
          });
          summary.cancelled++;
          continue;
        }

        if (sourceResolution.state === "failed") {
          throw new Error(sourceResolution.reason);
        }

        const recipient = String(rowValues[headers["recipient"] - 1] || "").trim();
        if (!recipient) throw new Error("Scheduled row is missing Recipient.");

        sentSheet.getRange(row, headers["status"]).setValue("Processing");
        SpreadsheetApp.flush();

        const templateKey = String(rowValues[headers["template"] - 1] || "").trim();
        const templateSheet = ss.getSheetByName("Templates");
        if (!templateSheet) throw new Error('Template sheet "Templates" no longer exists.');
        let template;
        try {
          template = getTemplateByKey_(templateSheet, templateKey);
        } catch (templateErr) {
          throw new Error('Template "' + templateKey + '" no longer exists.');
        }
        const senderProfile = getSenderProfile_(ss, sender);
        const rowData = getRowData_(
          sourceResolution.sheet,
          sourceResolution.row,
          sourceResolution.headers
        );
        const imageAssets = getImageAssets_(ss);
        validateLiveRenderingTokens_(template, senderProfile, rowData, imageAssets);
        const payload = buildEmailPayload_(
          ss,
          sourceResolution.sheet,
          sourceResolution.row,
          sourceResolution.headers,
          template,
          sender,
          { senderProfile: senderProfile, imageAssets: imageAssets, rowData: rowData }
        );
        renderedPayload = payload;

        const revalidatedEffectiveUser = String(getEffectiveUserEmail() || "").trim();
        if (
          !revalidatedEffectiveUser ||
          normalize_(revalidatedEffectiveUser) !== normalize_(sender)
        ) {
          throw new Error("Effective trigger owner no longer matches the queued sender.");
        }

        sendEmail({
          to: recipient,
          cc: String(rowValues[headers["cc"] - 1] || ""),
          bcc: String(rowValues[headers["bcc"] - 1] || ""),
          subject: payload.subject,
          body: payload.plainBody,
          htmlBody: payload.htmlBody,
          inlineImages: payload.inlineImages,
          attachments: payload.attachments,
          name: payload.senderName || sender
        });
        delivered = true;

        const successMessage = "Scheduled email sent successfully.";
        updateSentQueueRow_(sentSheet, row, rowValues, headers, {
          "Status": "Sent",
          "Processed At": new Date(now.getTime()),
          "Message": successMessage,
          "Name": metadata.recordIdValue,
          "Subject": payload.subject,
          "Email Body": payload.plainBody,
          "Attachments": payload.attachmentNames || "",
          "Log Note": successMessage
        });
        try {
          updateScheduledSourceTracker_(
            sourceResolution,
            templateKey,
            "Sent"
          );
        } catch (trackerErr) {
          Logger.log("SCHEDULED TRACKER UPDATE ERROR row " + row + ": " + trackerErr.message);
        }
        summary.sent++;
      } catch (err) {
        if (delivered) {
          Logger.log("SCHEDULED FINALIZATION ERROR row " + row + ": " + (err.message || err));
          continue;
        }

        const failureMessage = "Scheduled email failed: " + (err.message || err);
        const failureUpdates = {
          "Status": "Failed",
          "Processed At": new Date(now.getTime()),
          "Message": failureMessage,
          "Log Note": failureMessage
        };
        if (renderedPayload) {
          failureUpdates["Subject"] = renderedPayload.subject;
          failureUpdates["Email Body"] = renderedPayload.plainBody;
          failureUpdates["Attachments"] = renderedPayload.attachmentNames || "";
        }
        updateSentQueueRow_(sentSheet, row, rowValues, headers, failureUpdates);

        try {
          updateScheduledSourceTracker_(
            sourceResolution,
            rowValues[headers["template"] - 1],
            "Failed",
            err.message || String(err)
          );
        } catch (trackerErr) {
          Logger.log("SCHEDULED TRACKER UPDATE ERROR row " + row + ": " + trackerErr.message);
        }

        summary.failed++;
        Logger.log("SCHEDULED SEND ERROR row " + row + ": " + (err.message || err));
      }
    }

    return summary;
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// Email payload / variable rendering
// =============================================================================

function buildEmailPayload_(ss, trackerSheet, row, trackerHeaders, template, senderEmail, context) {
  const renderContext = context || {};
  const senderProfile = renderContext.senderProfile || getSenderProfile_(ss, senderEmail);
  const imageAssets = renderContext.imageAssets || getImageAssets_(ss);
  const inlineImages = {};
  const attachments = getAttachmentBlobsForTemplate_(template);

  const rowData = renderContext.rowData || getRowData_(trackerSheet, row, trackerHeaders);

  const textVars = Object.assign({}, rowData, {
    "Sender Name": senderProfile.name || "",
    "Sender Email": senderProfile.email || senderEmail || ""
  });

  const senderSignatureHtml = renderRichTextOrTextWithAssets_(
    senderProfile.signatureRichText,
    senderProfile.signatureText,
    textVars,
    imageAssets,
    inlineImages
  );

  const senderSignaturePlain = renderPlainTextWithAssets_(
    senderProfile.signatureText,
    textVars,
    imageAssets
  );

  const richHtmlVars = {
    "Sender Signature": senderSignatureHtml
  };

  const richPlainVars = {
    "Sender Signature": senderSignaturePlain
  };

  const subject = renderPlainTextWithAssets_(
    template.subject,
    textVars,
    imageAssets,
    richPlainVars
  );

  const htmlBody = renderTextWithAssetsToHtml_(
    template.body,
    textVars,
    imageAssets,
    inlineImages,
    richHtmlVars
  );

  const plainBody = renderPlainTextWithAssets_(
    template.body,
    textVars,
    imageAssets,
    richPlainVars
  );

  return {
    subject: subject,
    plainBody: plainBody,
    htmlBody: '<div style="' + getDefaultEmailWrapperStyle_() + '">' + htmlBody + '</div>',
    inlineImages: inlineImages,
    attachments: attachments,
    attachmentNames: attachments.map(blob => blob.getName()).join("\n"),
    senderName: senderProfile.name || senderEmail || ""
  };
}

function getAttachmentBlobsForTemplate_(template) {
  const raw = String(template.attachmentLink || "").trim();

  if (!raw) return [];

  const links = raw
    .split(/\n|,/)
    .map(value => value.trim())
    .filter(Boolean);

  return links.map(link => {
    const fileId = extractDriveFileId_(link);

    if (!fileId) {
      throw new Error("Could not extract Drive file ID from attachment link.");
    }

    try {
      return DriveApp.getFileById(fileId).getBlob();
    } catch (err) {
      throw new Error("Could not access attachment file. Check the Drive link and permissions.");
    }
  });
}

function renderRichTextOrTextWithAssets_(richTextValue, fallbackText, textVars, imageAssets, inlineImages) {
  if (!richTextValue) {
    return renderTextWithAssetsToHtml_(fallbackText, textVars, imageAssets, inlineImages, {});
  }

  const runs = richTextValue.getRuns();

  if (!runs || !runs.length) {
    return renderTextWithAssetsToHtml_(richTextValue.getText(), textVars, imageAssets, inlineImages, {});
  }

  let html = "";

  runs.forEach(run => {
    html += renderRichRunWithAssets_(run, textVars, imageAssets, inlineImages);
  });

  return html;
}

function renderRichRunWithAssets_(run, textVars, imageAssets, inlineImages) {
  const runText = run.getText();
  const style = run.getTextStyle();
  const linkUrl = run.getLinkUrl();

  const parts = String(runText || "").split(/(\{\{.*?\}\})/g);

  return parts.map(part => {
    const match = part.match(/^\{\{(.*?)\}\}$/);

    if (match) {
      const key = match[1].trim();

      if (Object.prototype.hasOwnProperty.call(textVars, key)) {
        return applyTextStyleHtml_(textToHtml_(textVars[key]), style, linkUrl);
      }

      const asset = imageAssets[normalizeTemplateKey_(key)];

      if (asset) {
        return getImageHtmlForAsset_(asset, inlineImages);
      }

      return "";
    }

    return applyTextStyleHtml_(textToHtml_(part), style, linkUrl);
  }).join("");
}

function applyTextStyleHtml_(htmlText, style, linkUrl) {
  if (!htmlText) return "";

  let css = "";

  if (style) {
    if (style.isBold()) css += "font-weight:700;";
    if (style.isItalic()) css += "font-style:italic;";
    if (style.isUnderline()) css += "text-decoration:underline;";

    const color = style.getForegroundColor();
    if (color) css += "color:" + color + ";";

    const fontSize = style.getFontSize();

    if (shouldPreserveRichTextFontSize_(fontSize)) {
      css += "font-size:" + fontSize + "pt;";
    }

    const fontFamily = style.getFontFamily();
    if (fontFamily) css += "font-family:" + escapeHtml_(fontFamily) + ", Arial, sans-serif;";
  }

  let output = css ? '<span style="' + css + '">' + htmlText + '</span>' : htmlText;

  if (linkUrl) {
    output = '<a href="' + escapeHtml_(linkUrl) + '" target="_blank">' + output + '</a>';
  }

  return output;
}
