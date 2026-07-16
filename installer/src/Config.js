const SENDMEBOT_INSTALLER_VERSION = "1.1.2";
const SENDMEBOT_RESERVED_SHEET_NAMES = [
  "Tracker", "Templates", "Senders", "Sent", "_SendMeBot"
];


function getInstallerEnvironment_() {
  if (
    typeof SENDMEBOT_INSTALLER_ENVIRONMENT === "undefined" ||
    !SENDMEBOT_INSTALLER_ENVIRONMENT
  ) {
    throw new Error("Installer environment was not generated. Run the installer build first.");
  }
  const config = SENDMEBOT_INSTALLER_ENVIRONMENT;
  if (!config.templateSpreadsheetId || config.templateSpreadsheetId === "REPLACE_LOCALLY") {
    throw new Error("The installer template Spreadsheet ID is not configured.");
  }
  return config;
}


function normalizeInstallerValue_(value) {
  return String(value || "").trim();
}


function getInstallerFormValue_(event, key) {
  const inputs = event && event.commonEventObject && event.commonEventObject.formInputs;
  const input = inputs && inputs[key];
  const values = input && input.stringInputs && input.stringInputs.value;
  return values && values.length ? String(values[0] || "") : "";
}


function getInstallerFormValues_(event, key) {
  const inputs = event && event.commonEventObject && event.commonEventObject.formInputs;
  const input = inputs && inputs[key];
  const values = input && input.stringInputs && input.stringInputs.value;
  return Array.isArray(values) ? values.map(String) : [];
}


function getInstallerActionParameter_(event, key) {
  const commonParameters = event && event.commonEventObject && event.commonEventObject.parameters;
  const legacyParameters = event && event.parameters;
  return normalizeInstallerValue_(
    commonParameters && commonParameters[key] !== undefined
      ? commonParameters[key]
      : legacyParameters && legacyParameters[key]
  );
}


function getInstallerSheetsContext_(event) {
  const sheets = event && event.sheets ? event.sheets : {};
  const eventHasFileScope = sheets.addonHasFileScopePermission === true;
  const actionSpreadsheetId = getInstallerActionParameter_(event, "sourceSpreadsheetId");
  const actionSpreadsheetName = getInstallerActionParameter_(event, "sourceSpreadsheetName");
  const spreadsheetId = eventHasFileScope
    ? normalizeInstallerValue_(sheets.id)
    : actionSpreadsheetId;
  return {
    hasFileScope: !!spreadsheetId,
    spreadsheetId: spreadsheetId,
    title: eventHasFileScope ? normalizeInstallerValue_(sheets.title) : actionSpreadsheetName
  };
}


function getCurrentInstallerSpreadsheet_(event) {
  const context = getInstallerSheetsContext_(event);
  if (!context.hasFileScope || !context.spreadsheetId) return null;
  try {
    return SpreadsheetApp.openById(context.spreadsheetId);
  } catch (err) {
    return null;
  }
}
