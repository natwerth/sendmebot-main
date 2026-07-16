const SENDMEBOT_UPDATE_TARGET_INDEX = "SENDMEBOT_UPDATE_TARGET_INDEX";
const SENDMEBOT_UPDATE_TARGET_PREFIX = "SENDMEBOT_UPDATE_TARGET_";
const SENDMEBOT_UPDATE_HANDLER = "runAutomaticUpdates";


function readSendMeBotRegistry_(ss) {
  const sheet = ss.getSheetByName("_SendMeBot");
  if (!sheet || sheet.getLastRow() < 1) return null;
  const result = {};
  sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues().forEach(row => {
    const key = String(row[0] || "").trim();
    if (key) result[key] = String(row[1] || "");
  });
  return result.appId === "sendmebot" ? result : null;
}


function makeUpdateTargetKey_(spreadsheetId) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(spreadsheetId || "")
  );
  return SENDMEBOT_UPDATE_TARGET_PREFIX + Utilities.base64EncodeWebSafe(digest).slice(0, 24);
}


function getUpdateTargetKeys_() {
  const raw = PropertiesService.getUserProperties().getProperty(SENDMEBOT_UPDATE_TARGET_INDEX);
  if (!raw) return [];
  try {
    const values = JSON.parse(raw);
    return Array.isArray(values) ? values.map(String) : [];
  } catch (err) {
    return [];
  }
}


function saveUpdateTarget_(target) {
  const properties = PropertiesService.getUserProperties();
  const key = makeUpdateTargetKey_(target.spreadsheetId);
  const keys = getUpdateTargetKeys_();
  if (keys.indexOf(key) === -1) keys.push(key);
  properties.setProperty(key, JSON.stringify(target));
  properties.setProperty(SENDMEBOT_UPDATE_TARGET_INDEX, JSON.stringify(keys));
  return key;
}


function getUpdateTargets_() {
  const properties = PropertiesService.getUserProperties();
  return getUpdateTargetKeys_().map(key => {
    try {
      return JSON.parse(properties.getProperty(key) || "null");
    } catch (err) {
      return null;
    }
  }).filter(Boolean);
}


function removeUpdateTarget_(spreadsheetId) {
  const properties = PropertiesService.getUserProperties();
  const key = makeUpdateTargetKey_(spreadsheetId);
  const keys = getUpdateTargetKeys_().filter(item => item !== key);
  properties.deleteProperty(key);
  properties.setProperty(SENDMEBOT_UPDATE_TARGET_INDEX, JSON.stringify(keys));
  if (!keys.length) removeAutomaticUpdateTriggers_();
}


function ensureAutomaticUpdateTrigger_() {
  const existing = ScriptApp.getProjectTriggers().filter(trigger =>
    trigger.getHandlerFunction() === SENDMEBOT_UPDATE_HANDLER
  );
  if (existing.length > 1) {
    existing.slice(1).forEach(trigger => ScriptApp.deleteTrigger(trigger));
  }
  if (existing.length) return existing[0];
  return ScriptApp.newTrigger(SENDMEBOT_UPDATE_HANDLER).timeBased().everyDays(1).create();
}


function removeAutomaticUpdateTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === SENDMEBOT_UPDATE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


function requestAppsScriptApi_(method, path, payload) {
  const options = {
    method: method,
    contentType: "application/json",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (payload !== undefined) options.payload = JSON.stringify(payload);
  const response = UrlFetchApp.fetch("https://script.googleapis.com/v1/" + path, options);
  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    if (status === 401 || status === 403) {
      throw new Error(
        "Apps Script API access is unavailable. Enable it at " +
        "https://script.google.com/home/usersettings and retry. HTTP " + status + "."
      );
    }
    throw new Error("Apps Script API request failed with HTTP " + status + ": " + text);
  }
  return text ? JSON.parse(text) : {};
}


function verifyUpdateTarget_(spreadsheetId, registry) {
  if (!registry || registry.appId !== "sendmebot" || !registry.scriptId) {
    throw new Error("This spreadsheet does not contain a valid SendMeBot installation registry.");
  }
  if (String(registry.spreadsheetId || "") !== String(spreadsheetId || "")) {
    throw new Error("The SendMeBot registry does not match the current spreadsheet.");
  }
  const project = requestAppsScriptApi_("get", "projects/" + encodeURIComponent(registry.scriptId));
  if (String(project.parentId || "") !== String(spreadsheetId || "")) {
    throw new Error("The registered Apps Script project is not bound to this spreadsheet.");
  }
  const content = requestAppsScriptApi_(
    "get",
    "projects/" + encodeURIComponent(registry.scriptId) + "/content"
  );
  const files = Array.isArray(content.files) ? content.files : [];
  const markerFound = files.some(file =>
    file.name === "Config" && /SENDMEBOT_APP_ID\s*=\s*["']sendmebot["']/.test(file.source || "")
  );
  if (!markerFound) throw new Error("The target project is not a recognized SendMeBot release.");
  return { project: project, files: files };
}


function enableAutomaticUpdates(e) {
  const ss = getCurrentInstallerSpreadsheet_(e);
  if (!ss) return notifyInstaller_("Open an installed SendMeBot spreadsheet first.");
  try {
    const registry = readSendMeBotRegistry_(ss);
    const verified = verifyUpdateTarget_(ss.getId(), registry);
    saveUpdateTarget_({
      spreadsheetId: ss.getId(),
      scriptId: registry.scriptId,
      enabledAt: new Date().toISOString(),
      installedVersion: registry.appVersion || "",
      lastContentHash: hashAppsScriptFiles_(verified.files),
      lastCheck: "",
      lastResult: "Enabled"
    });
    ensureAutomaticUpdateTrigger_();
    return notifyInstaller_("Automatic SendMeBot updates are enabled for this workbook.");
  } catch (err) {
    return notifyInstaller_(err.message || String(err));
  }
}


function disableAutomaticUpdates(e) {
  const ss = getCurrentInstallerSpreadsheet_(e);
  if (!ss) return notifyInstaller_("Open an installed SendMeBot spreadsheet first.");
  removeUpdateTarget_(ss.getId());
  return notifyInstaller_("Automatic updates are disabled for this workbook.");
}
