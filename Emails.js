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
    recipientFields: getRecipientFieldsForSendForm_(trackerSheet, headers, templates),
    scheduleOptions: [
      { value: "today", label: "Today" },
      { value: "start_minus_7", label: "7 days before Start Date" },
      { value: "custom", label: "Custom date" }
    ]
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
    scheduleOptions: [
      { value: "today", label: "Today" },
      { value: "start_minus_7", label: "7 days before Start Date" },
      { value: "custom", label: "Custom date" }
    ],
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
  const scheduleChoice = normalize_(formData.scheduleChoice || "today");
  const customDate = formData.customDate || "";

  if (!templateKey) throw new Error("Template is required.");
  if (!sender) throw new Error("Sender is required.");
  if (!toField) throw new Error("Recipient field is required.");

  if (action !== "send_now" && action !== "schedule") {
    throw new Error("Unknown action: " + formData.action);
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
    scheduleChoice: scheduleChoice,
    customDate: customDate,
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
    ensureScheduledSendTrigger_();

    const customDate = job.customDate ? parseUserDate_(job.customDate) : null;

    rows.forEach(row => {
      headers = getHeaders_(sheet);

      const result = scheduleOneRow_(
        ss,
        sheet,
        row,
        headers,
        job.template,
        job.sender,
        job.scheduleChoice,
        customDate,
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

function ensureScheduledSendTrigger_() {
  const functionName = "sendScheduledEmails";

  const alreadyExists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === functionName
  );

  if (alreadyExists) return;

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
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
// Schedule selected
// =============================================================================

function validateScheduledCell_(sheet, row, statusCol, scheduleData) {
  const cell = sheet.getRange(row, statusCol);
  const visibleValue = cell.getDisplayValue();

  if (!scheduleData || !scheduleData.scheduledDate) {
    return "Missing schedule metadata.";
  }

  if (!String(scheduleData.sender || "").trim()) {
    return "Missing sender in schedule metadata.";
  }

  const scheduledDate = new Date(scheduleData.scheduledDate);

  if (isNaN(scheduledDate.getTime())) {
    return "Invalid scheduledDate in metadata.";
  }

  const visibleDate = Utilities.formatDate(
    scheduledDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  if (visibleValue !== "Scheduled for " + visibleDate) {
    return (
      "Cell says '" +
      visibleValue +
      "' but metadata says Scheduled for " +
      visibleDate +
      "."
    );
  }

  return "";
}

function debugNote_(sheet, row, col, label) {
  const cell = sheet.getRange(row, col);
  Logger.log(
    label +
    " | " +
    sheet.getName() +
    "!" +
    cell.getA1Notation() +
    " | value=" +
    cell.getDisplayValue() +
    " | noteLength=" +
    String(cell.getNote() || "").length
  );
}

function scheduleOneRow_(
  ss,
  sheet,
  row,
  headers,
  templateKey,
  sender,
  scheduleChoice,
  customDate,
  recipientConfig
) {
  const statusCol = headers["status"];

  let recipients = {
    to: "",
    cc: "",
    bcc: ""
  };

  let email = "";
  const name = getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig);

  try {
    setStatus_(sheet, row, statusCol, "Scheduling...");

    recipients = resolveRecipientsForRow_(
      sheet,
      row,
      headers,
      recipientConfig
    );

    email = recipients.to;

    if (!email) throw new Error("Missing recipient email.");

    const scheduledDate = getScheduledDateForRow_(
      sheet,
      row,
      headers,
      scheduleChoice,
      customDate
    );

    const statusHeaderCol = getTemplateStatusColumn_(getHeaders_(sheet), templateKey, true);

    stampTemplateScheduled_(sheet, row, getHeaders_(sheet), templateKey, scheduledDate);
    debugNote_(sheet, row, statusHeaderCol, "After stampTemplateScheduled_");

    saveScheduledEmailData_(sheet, row, statusHeaderCol, {
      template: templateKey,
      sender: sender,
      scheduledDate: scheduledDate.toISOString(),
      recipientConfig: recipientConfig || {
        sender: sender,
        toField: { value: "Email", valueType: "field" },
        ccFields: [],
        bccFields: []
      }
    });
    debugNote_(sheet, row, statusHeaderCol, "After saveScheduledEmailData_");

    setStatus_(sheet, row, statusCol, "Scheduled");
    debugNote_(sheet, row, statusHeaderCol, "After setStatus_");

    logSentEmail_(ss, {
      name,
      email,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: templateKey,
      sender,
      subject: "",
      status: "Scheduled",
      message:
        "Email scheduled for " +
        Utilities.formatDate(scheduledDate, Session.getScriptTimeZone(), "M/d/yyyy"),
      sourceRow: row,
      body: "",
      attachments: ""
    });

    debugNote_(sheet, row, statusHeaderCol, "After logSentEmail_");

    return "scheduled";

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
      message: err.message || "Scheduling failed.",
      sourceRow: row,
      body: "",
      attachments: ""
    });

    Logger.log("SCHEDULE ERROR row " + row + ": " + err.message);
    return "failed";
  }
}

// =============================================================================
// Scheduled send trigger runner
// =============================================================================

function sendScheduledEmails() {
  return processScheduledEmails_();
}

function processScheduledEmails_(options) {
  const runtime = options || {};
  const lock = runtime.lock || LockService.getScriptLock();
  const summary = {
    sent: 0,
    failed: 0,
    skipped: 0,
    locked: false
  };

  if (!lock.tryLock(30000)) {
    summary.locked = true;
    Logger.log("SCHEDULED SEND SKIPPED: another scheduled send is already running.");
    return summary;
  }

  try {
    const ss = runtime.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = runtime.sheet || getTrackerSheet_();
    const templateSheet = runtime.templateSheet || ss.getSheetByName("Templates");

    if (!sheet) throw new Error("Missing Tracker sheet.");
    if (!templateSheet) throw new Error("Missing Templates sheet.");

    const headers = getHeaders_(sheet);
    const values = sheet.getDataRange().getValues();
    const statusColumns = getTemplateStatusColumns_(headers);
    const effectiveUser = Object.prototype.hasOwnProperty.call(runtime, "effectiveUser")
      ? String(runtime.effectiveUser || "")
      : Session.getEffectiveUser().getEmail();
    const sendEmail = runtime.sendEmail || sendScheduledEmail_;
    const today = runtime.now ? new Date(runtime.now) : new Date();

    today.setHours(0, 0, 0, 0);

    for (let r = 1; r < values.length; r++) {
      const row = r + 1;

      statusColumns.forEach(statusInfo => {
        const cellValue = values[r][statusInfo.col - 1];
        if (!isScheduledStatus_(cellValue)) return;

        try {
          processScheduledEmailCell_({
            ss: ss,
            sheet: sheet,
            templateSheet: templateSheet,
            headers: headers,
            row: row,
            statusInfo: statusInfo,
            today: today,
            effectiveUser: effectiveUser,
            sendEmail: sendEmail,
            summary: summary
          });
        } catch (err) {
          summary.failed++;
          Logger.log(
            "SCHEDULED CELL ERROR row " +
            row +
            ", column " +
            statusInfo.col +
            ": " +
            (err.message || err)
          );
        }
      });
    }

    ss.toast(
      "Scheduled send run complete. Sent: " +
        summary.sent +
        ". Failed: " +
        summary.failed +
        ". Skipped: " +
        summary.skipped +
        ".",
      "SendMeBot",
      8
    );

    return summary;
  } finally {
    lock.releaseLock();
  }
}

function processScheduledEmailCell_(context) {
  const sheet = context.sheet;
  const row = context.row;
  const statusInfo = context.statusInfo;
  const summary = context.summary;
  let scheduleData = null;
  let recipientConfig = null;
  let recipients = { to: "", cc: "", bcc: "" };
  let name = "";
  let email = "";

  try {
    scheduleData = getScheduledEmailData_(sheet, row, statusInfo.col);

    const scheduleProblem = validateScheduledCell_(
      sheet,
      row,
      statusInfo.col,
      scheduleData
    );

    if (scheduleProblem) {
      markScheduledCellBroken_(sheet, row, statusInfo.col, scheduleProblem);
      summary.skipped++;
      return;
    }

    const scheduledDate = new Date(scheduleData.scheduledDate);

    if (!isScheduledEmailDue_(scheduledDate, context.today)) return;

    if (normalize_(context.effectiveUser) !== normalize_(scheduleData.sender)) {
      summary.skipped++;
      Logger.log(
        "SCHEDULED SEND SKIPPED row " +
        row +
        ", column " +
        statusInfo.col +
        ": trigger owner does not match scheduled sender."
      );
      return;
    }

    recipientConfig = scheduleData.recipientConfig || {
      sender: scheduleData.sender || "",
      toField: { value: "Email", valueType: "field" },
      ccFields: [],
      bccFields: []
    };

    if (!recipientConfig.sender) {
      recipientConfig.sender = scheduleData.sender || "";
    }

    name = getRecipientNameForRow_(
      context.ss,
      sheet,
      row,
      context.headers,
      recipientConfig
    );

    recipients = resolveRecipientsForRow_(
      sheet,
      row,
      context.headers,
      recipientConfig
    );

    email = recipients.to;
    if (!email) throw new Error("Missing recipient email.");

    const statusCol = context.headers["status"];

    setStatus_(sheet, row, statusCol, "Sending scheduled email...");
    sheet.getRange(row, statusInfo.col).setValue("Sending scheduled email...");
    SpreadsheetApp.flush();

    const template = getTemplateByKey_(context.templateSheet, scheduleData.template);
    const payload = buildEmailPayload_(
      context.ss,
      sheet,
      row,
      context.headers,
      template,
      scheduleData.sender
    );

    context.sendEmail({
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

    stampTemplateColumn_(sheet, row, context.headers, scheduleData.template);
    deleteScheduledEmailData_(sheet, row, statusInfo.col);
    setStatus_(sheet, row, statusCol, "Sent");

    logSentEmail_(context.ss, {
      name: name,
      email: email,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: scheduleData.template,
      sender: scheduleData.sender,
      subject: payload.subject,
      status: "Sent",
      message: "Scheduled email sent successfully.",
      sourceRow: row,
      body: payload.plainBody,
      attachments: payload.attachmentNames || ""
    });

    summary.sent++;
  } catch (err) {
    if (!scheduleData) {
      markScheduledCellBroken_(
        sheet,
        row,
        statusInfo.col,
        err.message || "Scheduled email metadata could not be read."
      );
      summary.skipped++;
      return;
    }

    handleScheduledEmailFailure_(context, scheduleData, recipients, name, email, err);
    summary.failed++;
  }
}

function markScheduledCellBroken_(sheet, row, col, reason) {
  try {
    sheet.getRange(row, col).setValue("Schedule Broken");
  } catch (statusErr) {
    Logger.log(
      "SCHEDULE BROKEN STATUS ERROR row " +
      row +
      ", column " +
      col +
      ": " +
      statusErr.message
    );
  }

  Logger.log(
    "SCHEDULE BROKEN row " +
    row +
    ", column " +
    col +
    ": " +
    reason
  );
}

function handleScheduledEmailFailure_(context, scheduleData, recipients, name, email, err) {
  const row = context.row;

  try {
    stampTemplateFailure_(context.sheet, row, context.headers, scheduleData.template);
  } catch (statusErr) {
    Logger.log("SCHEDULED SEND TEMPLATE STATUS ERROR row " + row + ": " + statusErr.message);
  }

  try {
    setStatus_(context.sheet, row, context.headers["status"], "Failed");
  } catch (statusErr) {
    Logger.log("SCHEDULED SEND STATUS ERROR row " + row + ": " + statusErr.message);
  }

  try {
    logSentEmail_(context.ss, {
      name: name,
      email: email,
      cc: recipients.cc,
      bcc: recipients.bcc,
      template: scheduleData.template,
      sender: scheduleData.sender,
      subject: "",
      status: "Failed",
      message: err.message || "Scheduled email failed.",
      sourceRow: row,
      body: "",
      attachments: ""
    });
  } catch (logErr) {
    Logger.log("SCHEDULED SEND LOG ERROR row " + row + ": " + logErr.message);
  }

  Logger.log("SCHEDULED SEND ERROR row " + row + ": " + (err.message || err));
}

function sendScheduledEmail_(message) {
  MailApp.sendEmail(message);
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

// =============================================================================
// Date helpers
// =============================================================================

function isScheduledEmailDue_(scheduledDate, today) {
  const dueDate = new Date(scheduledDate);
  const currentDate = new Date(today);

  if (isNaN(dueDate.getTime())) throw new Error("Invalid scheduled date.");
  if (isNaN(currentDate.getTime())) throw new Error("Invalid current date.");

  dueDate.setHours(0, 0, 0, 0);
  currentDate.setHours(0, 0, 0, 0);

  return dueDate <= currentDate;
}

function getScheduledDateForRow_(sheet, row, headers, scheduleChoice, customDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (!scheduleChoice || scheduleChoice === "today") return today;

  if (scheduleChoice === "custom") {
    if (!(customDate instanceof Date)) throw new Error("Invalid custom date.");
    customDate.setHours(0, 0, 0, 0);
    return customDate;
  }

  if (scheduleChoice === "start_minus_7") {
    const startDateCol = headers["start date"];
    if (!startDateCol) throw new Error("Missing Start Date column.");

    const startDate = sheet.getRange(row, startDateCol).getValue();
    if (!(startDate instanceof Date)) throw new Error("Missing or invalid Start Date.");

    const scheduledDate = new Date(startDate);
    scheduledDate.setDate(scheduledDate.getDate() - 7);
    scheduledDate.setHours(0, 0, 0, 0);

    return scheduledDate;
  }

  throw new Error("Unknown schedule option: " + scheduleChoice);
}

function parseUserDate_(value) {
  const raw = String(value || "").trim();

  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);

    return new Date(year, month, day, 0, 0, 0, 0);
  }

  const parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return null;

  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

// =============================================================================
// Scheduled metadata
// =============================================================================

function saveScheduledEmailData_(sheet, row, statusCol, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const freshSheet = ss.getSheetByName(sheet.getName());
  const cell = freshSheet.getRange(row, statusCol);
  const json = JSON.stringify(data);

  cell.setNote(json);
  SpreadsheetApp.flush();

  Logger.log(
    "Immediately after setNote(): " +
    JSON.stringify(cell.getNote())
  );
}

function getScheduledEmailData_(sheet, row, statusCol) {
  const raw = sheet.getRange(row, statusCol).getNote();

  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error("Scheduled email metadata note is invalid JSON.");
  }
}

function deleteScheduledEmailData_(sheet, row, statusCol) {
  sheet.getRange(row, statusCol).clearNote();
}

function isScheduledStatus_(value) {
  return normalize_(value).startsWith("scheduled for ");
}
