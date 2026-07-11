function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SendMeBot")
    .addItem("Email selected now", "openEmailSelectedNowForm")
    .addItem("Schedule selected", "openScheduleSelectedForm")
    .addSeparator()
    .addItem("New template", "openComposeTemplateForm")
    .addItem("Edit template", "openEditTemplateForm")
    .addSeparator()
    .addItem("Sender profile", "openAddSenderForm")
    .addItem("Add image", "openAddImageForm")
    .addSeparator()
    .addItem("SendMeBot Assistant", "openSendMeBotAssistant")
    .addToUi();
}

function onEdit(e) {
  // Scheduling no longer stores metadata in tracker-cell notes. Retain this
  // simple-trigger entry point as a no-op for existing spreadsheet bindings.
  return;
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
    .setWidth(340)
    .setHeight(actionMode === "schedule" ? 650 : 545);

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
  let authenticatedEmail = "";
  let senderState = { status: "blocked", record: null, duplicateCount: 0 };
  let blockingError = "";

  try {
    authenticatedEmail = getAuthenticatedUserEmail_();
    senderState = toSenderBootstrapState_(
      getSenderStateForEmail_(ss, authenticatedEmail),
      true
    );
  } catch (err) {
    blockingError = err.message || String(err);
  }

  const template = HtmlService.createTemplateFromFile("AddSender");
  template.contextJson = JSON.stringify({
    imageAssets: sendersSheet ? getImageAssetsForComposer_(sendersSheet) : [],
    authenticatedEmail: authenticatedEmail,
    senderState: senderState,
    blockingError: blockingError
  });

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(435);

  SpreadsheetApp.getUi().showModelessDialog(html, "Sender profile");
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
