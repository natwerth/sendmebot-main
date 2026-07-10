function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SendMeBot")
    .addItem("Email selected now", "openEmailSelectedNowForm")
    .addItem("Schedule selected", "openScheduleSelectedForm")
    .addSeparator()
    .addItem("New template", "openComposeTemplateForm")
    .addItem("Edit template", "openEditTemplateForm")
    .addSeparator()
    .addItem("Add sender", "openAddSenderForm")
    .addItem("Add image", "openAddImageForm")
    .addSeparator()
    .addItem("SendMeBot Assistant", "openSendMeBotAssistant")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== getTrackerSheet_().getName()) return;

  const headers = getHeaders_(sheet);
  const statusColumns = getTemplateStatusColumns_(headers).map(info => info.col);

  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const values = range.getDisplayValues();

  let brokenCount = 0;
  let fakeScheduledCount = 0;

  for (let r = 0; r < numRows; r++) {
    const row = startRow + r;
    if (row === 1) continue;

    for (let c = 0; c < numCols; c++) {
      const col = startCol + c;

      if (statusColumns.indexOf(col) === -1) continue;

      const cellValue = values[r][c];
      const scheduleData = getScheduledEmailData_(sheet, row, col);

      if (scheduleData) {
        unscheduleCellFromManualEdit_(
          ss,
          sheet,
          row,
          col,
          headers,
          scheduleData,
          "Scheduled email was unscheduled because the tracking cell was manually edited."
        );

        brokenCount++;
        continue;
      }

      if (isScheduledStatus_(cellValue)) {
        const cell = sheet.getRange(row, col);

        cell.setValue("Error: Schedule Broken");
        SpreadsheetApp.flush();

        Utilities.sleep(1500);

        if (cell.getDisplayValue() === "Error: Schedule Broken") {
          cell.clearContent();
        }

        fakeScheduledCount++;

        Logger.log(
          "SCHEDULE BROKEN: scheduled text without metadata at row " +
          row +
          ", col " +
          col
        );
      }
    }
  }

  if (brokenCount || fakeScheduledCount) {
    let message = "";

    if (brokenCount && fakeScheduledCount) {
      message = "Some scheduled email cells were changed or entered without schedule data. Affected cells were cleared.";
    } else if (brokenCount) {
      message = "Scheduled email removed and logged in Sent.";
    } else {
      message = "Scheduled-looking text was entered without schedule data and cleared.";
    }

    ss.toast(message, "SendMeBot", 8);
  }
}

function unscheduleCellFromManualEdit_(ss, sheet, row, col, headers, scheduleData, reason) {
  const cell = sheet.getRange(row, col);

  const recipientConfig = scheduleData.recipientConfig || {
    sender: scheduleData.sender || "",
    toField: { value: "Email", valueType: "field" },
    ccFields: [],
    bccFields: []
  };

  let name = "";
  let recipients = { to: "", cc: "", bcc: "" };

  try {
    name = getRecipientNameForRow_(ss, sheet, row, headers, recipientConfig);
    recipients = resolveRecipientsForRow_(sheet, row, headers, recipientConfig);
  } catch (err) {
    Logger.log("UNSCHEDULE LOG RECIPIENT ERROR row " + row + ": " + err.message);
  }

  logSentEmail_(ss, {
    name: name,
    email: recipients.to,
    cc: recipients.cc,
    bcc: recipients.bcc,
    template: scheduleData.template || "",
    sender: scheduleData.sender || Session.getActiveUser().getEmail(),
    subject: "",
    status: "Unscheduled",
    message: reason || "Scheduled email was unscheduled by manual edit.",
    sourceRow: row,
    body: "",
    attachments: "",
    logNote: "Schedule metadata note was cleared from " + cell.getA1Notation() + "."
  });

  cell.setValue("Error: Schedule Broken");
  cell.clearNote();
  SpreadsheetApp.flush();

  Utilities.sleep(1500);

  if (cell.getDisplayValue() === "Error: Schedule Broken") {
    cell.clearContent();
  }
}

// Keeps old installable edit triggers from breaking.
function handleEdit(e) {
  return;
}


// --- Send form modal openers ---
function openEmailSelectedNowForm() {
  openSendMeBotForm_("send_now");
}


function openScheduleSelectedForm() {
  openSendMeBotForm_("schedule");
}


function openSendMeBotForm_(actionMode) {
  const template = HtmlService.createTemplateFromFile("SendForm");
  template.contextJson = JSON.stringify(getSendFormOpenContext_(actionMode));

  const html = template
    .evaluate()
    .setWidth(300)
    .setHeight(actionMode === "schedule" ? 315 : 235);

  const title = actionMode === "schedule"
    ? "Schedule selected"
    : "Email selected now";

  SpreadsheetApp.getUi().showModelessDialog(html, title);
}


// --- Template modal openers ---
function openComposeTemplateForm() {
  openTemplateForm_("new");
}

function openEditTemplateForm() {
  openTemplateForm_("edit");
}

function openTemplateForm_(mode) {
  const template = HtmlService.createTemplateFromFile("TemplateForm");
  template.contextJson = JSON.stringify(getTemplateFormContext_(mode));

  const html = template
    .evaluate()
    .setWidth(520)
    .setHeight(mode === "edit" ? 690 : 620);

  SpreadsheetApp.getUi().showModelessDialog(
    html,
    mode === "edit" ? "Edit template" : "New template"
  );
}

function openAddSenderForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sendersSheet = ss.getSheetByName("Senders");

  const template = HtmlService.createTemplateFromFile("AddSender");
  template.contextJson = JSON.stringify({
    imageAssets: sendersSheet ? getImageAssetsForComposer_(sendersSheet) : []
  });

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(435);

  SpreadsheetApp.getUi().showModelessDialog(html, "Add sender");
}


function openAddImageForm() {
  const html = HtmlService
    .createHtmlOutputFromFile("AddImage")
    .setWidth(400)
    .setHeight(275);

  SpreadsheetApp.getUi().showModelessDialog(html, "Add image");
}

// --- Assistant sidebar ---
function openSendMeBotAssistant() {
  const template = HtmlService.createTemplateFromFile("AssistantSidebar");
  template.gemUrl = "https://gemini.google.com/gem/6c4e76611b3f";

  const html = template
    .evaluate()
    .setTitle("SendMeBot Assistant");

  SpreadsheetApp.getUi().showSidebar(html);
}

const TRACKER_SHEET_NAME = "Hires & Conversion - Intern";

function getTrackerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRACKER_SHEET_NAME);

  if (!sheet) {
    throw new Error('Missing tracker sheet: "' + TRACKER_SHEET_NAME + '".');
  }

  return sheet;
}