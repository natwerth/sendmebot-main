const SENDMEBOT_INSTALL_OPERATION_PREFIX = "SENDMEBOT_INSTALL_OPERATION_";
const SENDMEBOT_INSTALL_OPERATION_TTL_MS = 2 * 60 * 1000;


function getInstallOperationKey_(kind, sourceId, name) {
  const raw = [kind, sourceId || "none", name || "SendMeBot"].join("|");
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return SENDMEBOT_INSTALL_OPERATION_PREFIX + Utilities.base64EncodeWebSafe(digest).slice(0, 24);
}


function getRecentInstallOperation_(key, now) {
  const raw = PropertiesService.getUserProperties().getProperty(key);
  if (!raw) return null;
  try {
    const operation = JSON.parse(raw);
    if (Number(now || Date.now()) - Number(operation.createdAt || 0) <= SENDMEBOT_INSTALL_OPERATION_TTL_MS) {
      return operation;
    }
  } catch (err) {
    return null;
  }
  return null;
}


function saveInstallOperation_(key, operation) {
  PropertiesService.getUserProperties().setProperty(key, JSON.stringify(operation));
}


function copySendMeBotTemplate_(name) {
  const environment = getInstallerEnvironment_();
  const response = Drive.Files.copy(
    { name: normalizeInstallerValue_(name) || "SendMeBot" },
    environment.templateSpreadsheetId
  );
  if (!response || !response.id) throw new Error("The template copy did not return a Spreadsheet ID.");
  return response.id;
}


function createNewSendMeBotWorkbook(e) {
  const name = normalizeInstallerValue_(getInstallerFormValue_(e, "newWorkbookName")) || "SendMeBot";
  return runInstallOperation_("new", "", name, function() {
    const spreadsheetId = copySendMeBotTemplate_(name);
    return {
      spreadsheetId: spreadsheetId,
      url: "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit",
      warnings: [],
      incomplete: false
    };
  });
}


function installIntoCurrentWorkbook(e) {
  const source = getCurrentInstallerSpreadsheet_(e);
  if (!source) return notifyInstaller_("Open the source spreadsheet and try again.");
  const selectedIds = getInstallerFormValues_(e, "sheetIds");
  if (!selectedIds.length) return notifyInstaller_("Select at least one sheet to copy.");
  const scan = scanWorkbookCompatibility_(source);
  const accepted = getInstallerFormValues_(e, "acceptComplexMigration").indexOf("accepted") !== -1;
  if (scan.warnings.length && !accepted) {
    return notifyInstaller_("Review the compatibility warnings and confirm before continuing.");
  }
  const destinationName = source.getName() + " — SendMeBot";

  return runInstallOperation_("migrate", source.getId(), destinationName, function() {
    const destinationId = copySendMeBotTemplate_(destinationName);
    const url = "https://docs.google.com/spreadsheets/d/" + destinationId + "/edit";
    try {
      const copied = copySelectedSheets_(source, destinationId, selectedIds);
      const warnings = auditCopiedSheets_(destinationId, copied);
      return {
        spreadsheetId: destinationId,
        url: url,
        warnings: warnings,
        incomplete: warnings.length > 0
      };
    } catch (err) {
      return {
        spreadsheetId: destinationId,
        url: url,
        warnings: ["Sheet migration stopped early: " + (err.message || err)],
        incomplete: true
      };
    }
  });
}


function runInstallOperation_(kind, sourceId, name, callback) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10000)) return notifyInstaller_("An installation is already running.");
  try {
    const key = getInstallOperationKey_(kind, sourceId, name);
    const previous = getRecentInstallOperation_(key, Date.now());
    if (previous && previous.result) return buildInstallCompleteResponse_(previous.result);
    const result = callback();
    saveInstallOperation_(key, { createdAt: Date.now(), result: result });
    return buildInstallCompleteResponse_(result);
  } catch (err) {
    return notifyInstaller_("Installation failed: " + (err.message || err));
  } finally {
    lock.releaseLock();
  }
}
