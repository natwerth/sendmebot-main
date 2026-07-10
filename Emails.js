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
  const sendersSheet = ss.getSheetByName("Senders");

  if (!trackerSheet) throw new Error("Missing Tracker sheet.");
  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const headers = getHeaders_(trackerSheet);
  const selectedRows = getSelectedRows_(trackerSheet, headers);
  const currentUser = Session.getActiveUser().getEmail();

  const templates = getTemplateKeys_(templateSheet);

  let senders = [];

  if (sendersSheet) {
    senders = getSenderEmails_(sendersSheet);
  }

  if (senders.indexOf(currentUser) === -1) {
    senders = [currentUser].concat(senders);
  }

  const senderOptions = getSenderOptionsForSendForm_(senders);

  return {
    actionMode: actionMode,
    selectedRowCount: selectedRows.length,
    currentUser: currentUser,
    templates: templates,
    senders: senders,
    senderOptions: senderOptions,
    recipientFields: getRecipientFieldsForSendForm_(trackerSheet, headers, templates)
  };
}

function getSenderOptionsForSendForm_(senderEmails) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sendersSheet = ss.getSheetByName("Senders");
  const senderNameMap = {};

  if (sendersSheet && sendersSheet.getLastRow() >= 2) {
    const values = sendersSheet
      .getRange(2, 1, sendersSheet.getLastRow() - 1, 2)
      .getValues();

    values.forEach(row => {
      const name = String(row[0] || "").trim();
      const email = String(row[1] || "").trim();

      if (email) {
        senderNameMap[normalize_(email)] = name || email;
      }
    });
  }

  return (senderEmails || [])
    .map(email => {
      const cleanEmail = String(email || "").trim();

      return {
        name: senderNameMap[normalize_(cleanEmail)] || cleanEmail,
        email: cleanEmail
      };
    })
    .filter(sender => sender.email);
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
  const sendersSheet = ss.getSheetByName("Senders");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const templates = getTemplateKeys_(templateSheet);

  let senders = [];

  if (sendersSheet) {
    senders = getSenderEmails_(sendersSheet);
  }

  return {
    templates: templates,
    senders: senders,
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


function isSendMeBotStatusColumn_(sheet, col, displayHeader) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const sampleSize = Math.min(lastRow - 1, 100);
  const values = sheet
    .getRange(2, 1, sampleSize, sheet.getLastColumn())
    .getDisplayValues();

  return isSendMeBotStatusColumnFromSample_(displayHeader, values, col);
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
  const sender = String(formData.sender || "").trim() || Session.getActiveUser().getEmail();
  const toField = formData.toField || { value: "Email", valueType: "field" };
  const ccFields = Array.isArray(formData.ccFields) ? formData.ccFields : [];
  const bccFields = Array.isArray(formData.bccFields) ? formData.bccFields : [];
  const scheduledForIso = String(formData.scheduledForIso || "").trim();
  const scheduledTimeZone = String(formData.scheduledTimeZone || "").trim();
  const scheduledDisplayText = String(formData.scheduledDisplayText || "").trim();

  if (!templateKey) throw new Error("Template is required.");
  if (!sender) throw new Error("Sender is required.");
  if (!toField) throw new Error("Recipient field is required.");

  if (action !== "send_now" && action !== "schedule") {
    throw new Error("Unknown action: " + formData.action);
  }

  if (action === "schedule") {
    validateScheduledInstant_(scheduledForIso, new Date());
    if (!scheduledTimeZone) throw new Error("Browser timezone is required for scheduling.");
    if (!scheduledDisplayText) throw new Error("Scheduled display time is required.");
  }

  ensureTrackerColumnForTemplate_(templateKey);
  headers = getHeaders_(sheet);

  const progressText = action === "schedule" ? "Scheduling..." : "Sending...";
  setTemplateProgressForRows_(sheet, selectedRows, headers, templateKey, progressText);

  const job = {
    jobId: makeJobId_(),
    createdAt: new Date().toISOString(),
    createdBy: Session.getActiveUser().getEmail(),
    action: action,
    template: templateKey,
    sender: sender,
    toField: toField,
    ccFields: ccFields,
    bccFields: bccFields,
    scheduledForIso: scheduledForIso,
    scheduledTimeZone: scheduledTimeZone,
    scheduledDisplayText: scheduledDisplayText,
    rows: selectedRows
  };

  saveQueuedJob_(job);

  // Current modeless flow: process immediately after handoff.
  // Backup menu/trigger can still process orphaned jobs.
  const result = processQueuedJobs() || {
    sent: 0,
    scheduled: 0,
    failed: 0,
    skipped: 0,
    locked: false
  };

  const successCount = (result.sent || 0) + (result.scheduled || 0);
  const failedCount = result.failed || 0;

  return {
    title: failedCount ? "Request completed with errors" : "Request completed",
    action: action,
    selected: selectedRows.length,
    sent: result.sent || 0,
    scheduled: result.scheduled || 0,
    failed: failedCount,
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
  const lock = LockService.getDocumentLock();

  const summary = {
    sent: 0,
    scheduled: 0,
    failed: 0,
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

        summary.sent += result.sent || 0;
        summary.scheduled += result.scheduled || 0;
        summary.failed += result.failed || 0;
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
  const templateSheet = ss.getSheetByName("Templates");

  if (!sheet) throw new Error("Missing Tracker sheet.");
  if (!templateSheet) throw new Error("Missing Templates sheet.");

  let headers = getHeaders_(sheet);
  const rows = job.rows || [];
  const recipientConfig = getRecipientConfigFromJob_(job);

  let sentCount = 0;
  let scheduledCount = 0;
  let failedCount = 0;
  let skippedCount = 0;

  if (job.action === "send_now") {
    rows.forEach(row => {
      const result = sendOneRowNow_(
        ss,
        sheet,
        row,
        headers,
        job.template,
        job.sender,
        recipientConfig
      );

      result === "sent" ? sentCount++ : failedCount++;
    });

    clearSelectedRows_(sheet, headers, rows);

    ss.toast(
      "Send complete. Sent: " + sentCount + ". Failed: " + failedCount + ".",
      "SendMeBot",
      8
    );

    return {
      sent: sentCount,
      scheduled: 0,
      failed: failedCount,
      skipped: skippedCount
    };
  }

  if (job.action === "schedule") {
    const scheduledDate = parseScheduledInstant_(job.scheduledForIso);

    rows.forEach(row => {
      headers = getHeaders_(sheet);

      const result = scheduleOneRow_(
        ss,
        sheet,
        row,
        headers,
        job.template,
        job.sender,
        scheduledDate,
        job.scheduledTimeZone,
        job.scheduledDisplayText,
        recipientConfig
      );

      result === "scheduled" ? scheduledCount++ : failedCount++;
    });

    clearSelectedRows_(sheet, headers, rows);

    ss.toast(
      "Scheduling complete. Scheduled: " + scheduledCount + ". Failed: " + failedCount + ".",
      "SendMeBot",
      8
    );

    return {
      sent: 0,
      scheduled: scheduledCount,
      failed: failedCount,
      skipped: skippedCount
    };
  }

  throw new Error("Unknown queued action: " + job.action);
}

function setupScheduledEmailTrigger() {
  const functionName = "sendScheduledEmails";

  const alreadyExists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === functionName
  );

  if (alreadyExists) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Scheduled email processor trigger is already set up.",
      "SendMeBot",
      5
    );
    return;
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Scheduled emails will be checked every 5 minutes.",
    "SendMeBot",
    5
  );
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

function sendOneRowNow_(ss, sheet, row, headers, templateKey, sender, recipientConfig) {
  const currentUser = Session.getActiveUser().getEmail();
  const statusCol = headers["status"];
  const templateSheet = ss.getSheetByName("Templates");

  let recipients = {
    to: "",
    cc: "",
    bcc: ""
  };

  let email = "";

  const name = getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig);

  try {
    setStatus_(sheet, row, statusCol, "Sending...");

    recipients = resolveRecipientsForRow_(
      sheet,
      row,
      headers,
      recipientConfig
    );

    email = recipients.to;

    if (!email) throw new Error("Missing recipient email.");

    if (normalize_(currentUser) !== normalize_(sender)) {
      throw new Error("Failed: Sender does not match the active user. Ask Nat for help.");
    }

    const template = getTemplateByKey_(templateSheet, templateKey);
    const payload = buildEmailPayload_(ss, sheet, row, headers, template, sender);

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

    stampTemplateColumn_(sheet, row, getHeaders_(sheet), templateKey);
    setStatus_(sheet, row, statusCol, "Sent");

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

    return "sent";

  } catch (err) {
    stampTemplateFailure_(sheet, row, getHeaders_(sheet), templateKey);
    setStatus_(sheet, row, statusCol, "Failed");

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

    Logger.log("SEND ERROR row " + row + ": " + err.message);
    return "failed";
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

function getFrozenDeliveryMetadata_(ss, sheet, row, statusColumn, template, payload, audit) {
  const imageAssets = getImageAssets_(ss);
  const inlineImageFileIds = {};

  Object.keys(payload.inlineImages || {}).forEach(cid => {
    const matchingAsset = Object.keys(imageAssets)
      .map(key => imageAssets[key])
      .find(asset => getAssetCid_(asset.name) === cid);
    const fileId = matchingAsset ? extractDriveFileId_(matchingAsset.link) : "";
    if (!fileId) throw new Error("Could not preserve an inline image for scheduled delivery.");
    inlineImageFileIds[cid] = fileId;
  });

  const attachmentLinks = String(template.attachmentLink || "")
    .split(/\n|,/)
    .map(value => value.trim())
    .filter(Boolean);

  const attachmentFileIds = attachmentLinks.map(link => {
    const fileId = extractDriveFileId_(link);
    if (!fileId) throw new Error("Could not preserve an attachment for scheduled delivery.");
    return fileId;
  });

  return {
    version: 2,
    sourceSheet: sheet.getName(),
    sourceRow: row,
    sourceStatusColumn: statusColumn,
    scheduledTimeZone: audit.scheduledTimeZone,
    scheduledDisplayText: audit.scheduledDisplayText,
    senderName: payload.senderName || "",
    htmlBody: payload.htmlBody || "",
    inlineImageFileIds: inlineImageFileIds,
    attachmentFileIds: attachmentFileIds,
    attachmentNames: payload.attachmentNames || ""
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
  recipientConfig
) {
  const statusCol = headers["status"];
  let queueRowWritten = false;

  try {
    if (!(scheduledDate instanceof Date) || isNaN(scheduledDate.getTime())) {
      throw new Error("Invalid scheduled datetime.");
    }

    if (normalize_(Session.getActiveUser().getEmail()) !== normalize_(sender)) {
      throw new Error("Failed: Sender does not match the active user. Ask Nat for help.");
    }

    const recipients = resolveRecipientsForRow_(sheet, row, headers, recipientConfig);
    const name = getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig);
    if (!recipients.to) throw new Error("Missing recipient email.");

    const templateSheet = ss.getSheetByName("Templates");
    if (!templateSheet) throw new Error("Missing Templates sheet.");

    const template = getTemplateByKey_(templateSheet, templateKey);
    const payload = buildEmailPayload_(ss, sheet, row, headers, template, sender);
    const statusHeaderCol = getTemplateStatusColumn_(getHeaders_(sheet), templateKey, true);
    const metadata = getFrozenDeliveryMetadata_(
      ss,
      sheet,
      row,
      statusHeaderCol,
      template,
      payload,
      {
        scheduledTimeZone: scheduledTimeZone,
        scheduledDisplayText: scheduledDisplayText
      }
    );

    logSentEmail_(ss, {
      timestamp: new Date(),
      status: "Scheduled",
      scheduledFor: new Date(scheduledDate.getTime()),
      processedAt: "",
      message: "Email scheduled for " + scheduledDisplayText + ".",
      name: name,
      email: recipients.to,
      sender: sender,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: templateKey,
      subject: payload.subject,
      body: payload.plainBody,
      attachments: payload.attachmentNames || "",
      logNote: JSON.stringify(metadata)
    });
    queueRowWritten = true;

    const visibleScheduledDate = Utilities.formatDate(
      scheduledDate,
      ss.getSpreadsheetTimeZone(),
      "M/d"
    );
    sheet.getRange(row, statusHeaderCol).setValue("Scheduled for " + visibleScheduledDate);
    setStatus_(sheet, row, statusCol, "Scheduled");
    return "scheduled";
  } catch (err) {
    if (!queueRowWritten) {
      stampTemplateFailure_(sheet, row, getHeaders_(sheet), templateKey);
      setStatus_(sheet, row, statusCol, "Failed");
    } else {
      Logger.log(
        "SCHEDULE TRACKER UPDATE ERROR row " + row +
        ": queue row was written and will still be processed. " + (err.message || err)
      );
      return "scheduled";
    }

    Logger.log("SCHEDULE ERROR row " + row + ": " + (err.message || err));
    return "failed";
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

  if (!metadata || metadata.version !== 2) {
    throw new Error("Scheduled row has unsupported delivery metadata.");
  }

  return metadata;
}

function getScheduledBlobs_(metadata, getFileBlob) {
  const loadBlob = getFileBlob || function(fileId) {
    return DriveApp.getFileById(fileId).getBlob();
  };
  const inlineImages = {};

  Object.keys(metadata.inlineImageFileIds || {}).forEach(cid => {
    inlineImages[cid] = loadBlob(metadata.inlineImageFileIds[cid]);
  });

  return {
    inlineImages: inlineImages,
    attachments: (metadata.attachmentFileIds || []).map(loadBlob)
  };
}

function updateSentQueueRow_(sheet, row, rowValues, headers, updates) {
  Object.keys(updates).forEach(header => {
    const col = headers[normalize_(header)];
    if (col) rowValues[col - 1] = updates[header];
  });
  sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
}

function validateScheduledSourceTracker_(ss, metadata, templateKey) {
  const sourceSheet = ss.getSheetByName(String(metadata.sourceSheet || ""));
  const sourceRow = Number(metadata.sourceRow);
  const sourceCol = Number(metadata.sourceStatusColumn);

  if (!sourceSheet) {
    return { state: "orphaned", reason: "source sheet no longer exists" };
  }

  if (!Number.isInteger(sourceRow) || sourceRow < 2 || sourceRow > sourceSheet.getLastRow()) {
    return { state: "orphaned", reason: "source row no longer exists" };
  }

  if (!Number.isInteger(sourceCol) || sourceCol < 1 || sourceCol > sourceSheet.getLastColumn()) {
    return { state: "orphaned", reason: "source status column no longer exists" };
  }

  const currentHeaders = getHeaders_(sourceSheet);
  if (getTemplateStatusColumn_(currentHeaders, templateKey, false) !== sourceCol) {
    return { state: "orphaned", reason: "template-specific source status column no longer exists" };
  }

  const currentValue = sourceSheet.getRange(sourceRow, sourceCol).getDisplayValue();
  if (currentValue.indexOf("Scheduled for ") !== 0) {
    return { state: "cancelled", reason: "tracker status was cleared or changed" };
  }

  return { state: "valid" };
}

function updateScheduledSourceTracker_(ss, metadata, templateKey, outcome) {
  const sourceSheet = ss.getSheetByName(String(metadata.sourceSheet || ""));
  const sourceRow = Number(metadata.sourceRow);
  const sourceCol = Number(metadata.sourceStatusColumn);

  if (!sourceSheet || sourceRow < 2 || sourceRow > sourceSheet.getLastRow()) return false;
  if (sourceCol < 1 || sourceCol > sourceSheet.getLastColumn()) return false;

  const current = sourceSheet.getRange(sourceRow, sourceCol).getDisplayValue();
  const currentHeaders = getHeaders_(sourceSheet);
  if (getTemplateStatusColumn_(currentHeaders, templateKey, false) !== sourceCol) return false;
  if (current.indexOf("Scheduled for ") !== 0) return false;

  const trackerHeaders = currentHeaders;
  if (outcome === "Sent") {
    stampTemplateColumn_(sourceSheet, sourceRow, trackerHeaders, templateKey);
    setStatus_(sourceSheet, sourceRow, trackerHeaders["status"], "Sent");
  } else {
    stampTemplateFailure_(sourceSheet, sourceRow, trackerHeaders, templateKey);
    setStatus_(sourceSheet, sourceRow, trackerHeaders["status"], "Failed");
  }
  return true;
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
    const effectiveUser = Object.prototype.hasOwnProperty.call(runtime, "effectiveUser")
      ? String(runtime.effectiveUser || "")
      : Session.getEffectiveUser().getEmail();
    const sendEmail = runtime.sendEmail || function(message) { MailApp.sendEmail(message); };

    for (let r = 1; r < values.length; r++) {
      const rowValues = values[r].slice();
      const row = r + 1;
      if (rowValues[headers["status"] - 1] !== "Scheduled") continue;

      const scheduledFor = rowValues[headers["scheduled for"] - 1];

      if (!(scheduledFor instanceof Date) || isNaN(scheduledFor.getTime())) {
        const invalidMessage = "Scheduled For must be a valid spreadsheet datetime.";
        updateSentQueueRow_(sentSheet, row, rowValues, headers, {
          "Status": "Failed",
          "Processed At": new Date(now.getTime()),
          "Message": invalidMessage
        });
        summary.failed++;
        continue;
      }

      if (scheduledFor.getTime() > now.getTime()) {
        summary.skipped++;
        continue;
      }

      let delivered = false;

      try {
        const metadata = parseScheduledLogNote_(rowValues[headers["log note"] - 1]);
        const sourceValidation = validateScheduledSourceTracker_(
          ss,
          metadata,
          rowValues[headers["template"] - 1]
        );

        if (sourceValidation.state === "orphaned") {
          updateSentQueueRow_(sentSheet, row, rowValues, headers, {
            "Status": "Orphaned",
            "Processed At": new Date(now.getTime()),
            "Message": "Scheduled email was not sent because its " + sourceValidation.reason + "."
          });
          summary.orphaned++;
          continue;
        }

        if (sourceValidation.state === "cancelled") {
          updateSentQueueRow_(sentSheet, row, rowValues, headers, {
            "Status": "Cancelled",
            "Processed At": new Date(now.getTime()),
            "Message": "Scheduled email was cancelled because the tracker status was cleared or changed."
          });
          summary.cancelled++;
          continue;
        }

        const sender = String(rowValues[headers["sender"] - 1] || "").trim();
        const recipient = String(rowValues[headers["recipient"] - 1] || "").trim();
        if (!recipient) throw new Error("Scheduled row is missing Recipient.");
        if (!sender) throw new Error("Scheduled row is missing Sender.");
        if (normalize_(effectiveUser) !== normalize_(sender)) {
          throw new Error("Trigger owner does not match the scheduled sender.");
        }

        sentSheet.getRange(row, headers["status"]).setValue("Processing");
        SpreadsheetApp.flush();

        const blobs = getScheduledBlobs_(metadata, runtime.getFileBlob);

        sendEmail({
          to: recipient,
          cc: String(rowValues[headers["cc"] - 1] || ""),
          bcc: String(rowValues[headers["bcc"] - 1] || ""),
          subject: String(rowValues[headers["subject"] - 1] || ""),
          body: String(rowValues[headers["email body"] - 1] || ""),
          htmlBody: String(metadata.htmlBody || ""),
          inlineImages: blobs.inlineImages,
          attachments: blobs.attachments,
          name: String(metadata.senderName || sender)
        });
        delivered = true;

        updateSentQueueRow_(sentSheet, row, rowValues, headers, {
          "Status": "Sent",
          "Processed At": new Date(now.getTime()),
          "Message": "Scheduled email sent successfully."
        });
        try {
          updateScheduledSourceTracker_(
            ss,
            metadata,
            rowValues[headers["template"] - 1],
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
        updateSentQueueRow_(sentSheet, row, rowValues, headers, {
          "Status": "Failed",
          "Processed At": new Date(now.getTime()),
          "Message": failureMessage
        });

        try {
          const metadata = parseScheduledLogNote_(rowValues[headers["log note"] - 1]);
          updateScheduledSourceTracker_(
            ss,
            metadata,
            rowValues[headers["template"] - 1],
            "Failed"
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

// Manual migration only. It never sends mail or creates/changes triggers.
function migrateLegacyScheduledNotesToSentQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTrackerSheet_();
  const templateSheet = ss.getSheetByName("Templates");
  if (!sheet) throw new Error("Missing Tracker sheet.");
  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const headers = getHeaders_(sheet);
  const values = sheet.getDataRange().getDisplayValues();
  const notes = sheet.getDataRange().getNotes();
  const statusColumns = getTemplateStatusColumns_(headers);
  const report = { migrated: 0, invalid: 0, details: [] };

  for (let r = 1; r < values.length; r++) {
    statusColumns.forEach(statusInfo => {
      const rawNote = String(notes[r][statusInfo.col - 1] || "").trim();
      if (!rawNote) return;

      const location = sheet.getName() + "!" + sheet.getRange(r + 1, statusInfo.col).getA1Notation();

      try {
        const legacy = JSON.parse(rawNote);
        const scheduledDate = new Date(legacy.scheduledDate);
        if (!legacy.template || !legacy.sender || isNaN(scheduledDate.getTime())) {
          throw new Error("missing template, sender, or valid scheduledDate");
        }

        const expectedVisible = "Scheduled for " + Utilities.formatDate(
          scheduledDate,
          Session.getScriptTimeZone(),
          "M/d"
        );
        if (values[r][statusInfo.col - 1] !== expectedVisible) {
          throw new Error("visible schedule does not match note metadata");
        }
        if (getTemplateStatusColumn_(headers, legacy.template, false) !== statusInfo.col) {
          throw new Error("template does not match the noted status column");
        }
        if (normalize_(legacy.sender) !== normalize_(Session.getActiveUser().getEmail())) {
          throw new Error("scheduled sender does not match the active user");
        }

        const recipientConfig = legacy.recipientConfig || {
          sender: legacy.sender,
          toField: { value: "Email", valueType: "field" },
          ccFields: [],
          bccFields: []
        };
        const recipients = resolveRecipientsForRow_(sheet, r + 1, headers, recipientConfig);
        if (!recipients.to) throw new Error("missing recipient email");

        const template = getTemplateByKey_(templateSheet, legacy.template);
        const payload = buildEmailPayload_(ss, sheet, r + 1, headers, template, legacy.sender);
        const displayText = Utilities.formatDate(
          scheduledDate,
          Session.getScriptTimeZone(),
          "M/d/yyyy h:mm a z"
        );
        const metadata = getFrozenDeliveryMetadata_(
          ss,
          sheet,
          r + 1,
          statusInfo.col,
          template,
          payload,
          {
            scheduledTimeZone: Session.getScriptTimeZone(),
            scheduledDisplayText: displayText
          }
        );

        logSentEmail_(ss, {
          timestamp: new Date(),
          status: "Scheduled",
          scheduledFor: scheduledDate,
          processedAt: "",
          message: "Migrated legacy schedule for " + displayText + ".",
          name: getRecipientNameForRow_(ss, sheet, r + 1, headers, recipientConfig),
          email: recipients.to,
          sender: legacy.sender,
          cc: recipients.cc,
          bcc: recipients.bcc,
          template: legacy.template,
          subject: payload.subject,
          body: payload.plainBody,
          attachments: payload.attachmentNames || "",
          logNote: JSON.stringify(metadata)
        });

        sheet.getRange(r + 1, statusInfo.col).clearNote();
        report.migrated++;
      } catch (err) {
        report.invalid++;
        report.details.push(location + ": " + (err.message || err));
      }
    });
  }

  report.details.forEach(detail => Logger.log("LEGACY SCHEDULE NOT MIGRATED: " + detail));
  ss.toast(
    "Legacy schedules migrated: " + report.migrated + ". Invalid/ambiguous: " + report.invalid + ".",
    "SendMeBot",
    10
  );
  return report;
}

// =============================================================================
// Email payload / variable rendering
// =============================================================================

function buildEmailPayload_(ss, trackerSheet, row, trackerHeaders, template, senderEmail) {
  const senderProfile = getSenderProfile_(ss, senderEmail);
  const imageAssets = getImageAssets_(ss);
  const inlineImages = {};
  const attachments = getAttachmentBlobsForTemplate_(template);

  const rowData = getRowData_(trackerSheet, row, trackerHeaders);

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
