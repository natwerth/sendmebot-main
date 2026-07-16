function onOpen() {
  let registry = null;
  try {
    registry = ensureSendMeBotInstallationRegistry_();
  } catch (registryErr) {
    Logger.log("INSTALLATION REGISTRY ERROR: " + (registryErr.message || registryErr));
  }

  const environment = getSendMeBotEnvironmentConfig_();
  const onboardingPending = !!(registry && registry.onboardingState === "pending");
  const menu = SpreadsheetApp.getUi().createMenu(environment.brandName);

  if (onboardingPending) {
    menu.addItem("Continue setup", "openSendMeBotWalkthrough").addSeparator();
  }

  menu.addItem("Email selected now", "openEmailSelectedNowForm")
    .addItem("Schedule selected", "openScheduleSelectedForm")
    .addSeparator()
    .addItem("New template", "openComposeTemplateForm")
    .addItem("Edit template", "openEditTemplateForm")
    .addSeparator()
    .addItem("Sender profile", "openAddSenderForm")
    .addItem("Add image", "openAddImageForm")
    .addSeparator()
    .addItem("Setup", "openSetupForm");

  if (environment.assistantUrl) {
    menu.addSeparator().addItem(environment.brandName + " Assistant", "openSendMeBotAssistant");
  }

  menu.addToUi();

  if (onboardingPending && registry.onboardingAutoPrompted !== "true") {
    try {
      updateSendMeBotInstallationRegistry_({ onboardingAutoPrompted: "true" });
      openSendMeBotWalkthrough();
    } catch (walkthroughErr) {
      Logger.log("WALKTHROUGH PROMPT ERROR: " + (walkthroughErr.message || walkthroughErr));
    }
  }
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
    .setHeight(actionMode === "schedule" ? 550 : 375);

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
    blockingError: blockingError,
    brand: getSendMeBotClientBrand_()
  });

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(330);

  SpreadsheetApp.getUi().showModelessDialog(html, "Sender profile");
}


function openAddImageForm() {
  const html = HtmlService
    .createTemplateFromFile("AddImage")
    .evaluate()
    .setWidth(400)
    .setHeight(275);

  SpreadsheetApp.getUi().showModelessDialog(html, "Add image");
}

function includeHtml_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openSetupForm() {
  const template = HtmlService.createTemplateFromFile("SetupForm");
  template.contextJson = JSON.stringify(getSendMeBotSetupContext_());
  const html = template.evaluate().setWidth(340).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, "SendMeBot setup");
}


function openSendMeBotWalkthrough() {
  const html = HtmlService
    .createTemplateFromFile("WalkthroughForm")
    .evaluate()
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Welcome to SendMeBot");
}

// --- Assistant sidebar ---
function openSendMeBotAssistant() {
  const environment = getSendMeBotEnvironmentConfig_();
  const template = HtmlService.createTemplateFromFile("AssistantSidebar");
  template.gemUrl = environment.assistantUrl;
  template.logoUrl = environment.logoUrl;
  template.brandName = environment.brandName;

  const html = template
    .evaluate()
    .setTitle("SendMeBot Assistant");

  SpreadsheetApp.getUi().showSidebar(html);
}

const TRACKER_SHEET_PROPERTY = "SENDMEBOT_TRACKER_SHEET";
const RECORD_ID_HEADER_PROPERTY = "SENDMEBOT_RECORD_ID_HEADER";

function getSendMeBotConfig_(runtime) {
  const options = runtime || {};
  const properties = options.properties || PropertiesService.getDocumentProperties();
  const environment = getSendMeBotEnvironmentConfig_(options);
  return {
    trackerSheetName: String(
      properties.getProperty(TRACKER_SHEET_PROPERTY) || environment.defaultTrackerSheetName
    ),
    recordIdHeader: String(
      properties.getProperty(RECORD_ID_HEADER_PROPERTY) || environment.defaultRecordIdHeader
    )
  };
}

function getSendMeBotSetupContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSendMeBotConfig_();
  return {
    config: config,
    sheets: ss.getSheets()
      .filter(sheet => sheet.getName() !== SENDMEBOT_INTERNAL_SHEET)
      .map(sheet => {
        const headers = sheet.getLastColumn() > 0
          ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
              .map(value => String(value || "").trim()).filter(Boolean)
          : [];
        return {
          name: sheet.getName(),
          headers: headers,
          hasSelectColumn: headers.some(value => normalize_(value) === "select"),
          emailHeaders: headers.filter(value => /email/i.test(value))
        };
      }),
    brand: getSendMeBotClientBrand_()
  };
}

function saveSendMeBotSetup(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheetName = String((formData && formData.trackerSheetName) || "").trim();
  const recordIdHeader = String((formData && formData.recordIdHeader) || "").trim();
  const addSelectColumn = !!(formData && formData.addSelectColumn);
  const sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) throw new Error("Selected tracker sheet no longer exists.");
  if (!recordIdHeader) throw new Error("Record ID column is required.");

  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(value => String(value || "").trim());
  const normalizedHeaders = rawHeaders.filter(Boolean).map(normalize_);
  const duplicateHeaders = normalizedHeaders.filter((value, index) =>
    normalizedHeaders.indexOf(value) !== index
  );
  if (duplicateHeaders.length) {
    throw new Error("Tracker headers must be unique before Setup can be saved.");
  }

  let headers = getHeaders_(sheet);
  if (!headers["select"] && addSelectColumn) {
    addSelectColumnToTracker_(sheet);
    headers = getHeaders_(sheet);
  }
  if (!headers["select"]) {
    throw new Error('Selected tracker sheet is missing the "Select" column.');
  }
  if (!headers[normalize_(recordIdHeader)]) {
    throw new Error('Record ID column "' + recordIdHeader + '" was not found.');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    PropertiesService.getDocumentProperties().setProperties({
      [TRACKER_SHEET_PROPERTY]: trackerSheetName,
      [RECORD_ID_HEADER_PROPERTY]: recordIdHeader,
      [SENDMEBOT_SETUP_VERSION_PROPERTY]: SENDMEBOT_SETUP_VERSION
    });
    ensureSendMeBotInstallationRegistry_(ss);
  } finally {
    lock.releaseLock();
  }
  return { trackerSheetName: trackerSheetName, recordIdHeader: recordIdHeader };
}

function getTrackerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSendMeBotConfig_();
  const sheet = ss.getSheetByName(config.trackerSheetName);

  if (!sheet) {
    throw new Error('Missing tracker sheet: "' + config.trackerSheetName + '". Open SendMeBot → Setup.');
  }

  return sheet;
}
