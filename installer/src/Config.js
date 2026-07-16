const SENDMEBOT_INSTALLER_VERSION = "1.0.0";
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


function getCurrentInstallerSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    return null;
  }
}
