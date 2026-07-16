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
    const result = {
      spreadsheetId: spreadsheetId,
      url: "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit",
      warnings: [],
      incomplete: false
    };
    try {
      markSendMeBotOnboarding_(spreadsheetId, {
        installMode: "new",
        sourceSpreadsheetName: "",
        importedSheets: [],
        suggestedTrackerSheet: "Tracker"
      });
    } catch (err) {
      result.warnings.push("The workbook was created, but guided setup could not be initialized: " + (err.message || err));
      result.incomplete = true;
    }
    return result;
  });
}


function getInstallerDestinationName_(e, source) {
  const requestedName = normalizeInstallerValue_(
    getInstallerFormValue_(e, "destinationWorkbookName") ||
    getInstallerFormValue_(e, "newWorkbookName")
  );
  return requestedName || source.getName() + " — SendMeBot";
}


function getAllInstallerSheetIds_(scan) {
  return (scan && scan.sheets ? scan.sheets : []).map(sheet => String(sheet.sheetId));
}


function installEntireCurrentWorkbook(e) {
  const source = getCurrentInstallerSpreadsheet_(e);
  if (!source) return notifyInstaller_("Open the source spreadsheet and try again.");
  const scan = scanWorkbookCompatibility_(source);
  if (!scan.sheets.length) return notifyInstaller_("No worksheet tabs are available to copy.");
  const accepted = getInstallerFormValues_(e, "acceptComplexMigration").indexOf("accepted") !== -1;
  if (scan.warnings.length && !accepted) {
    return buildEntireWorkbookWarningResponse_(e, source, scan);
  }
  return installCurrentWorkbookSheets_(e, source, getAllInstallerSheetIds_(scan));
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
  return installCurrentWorkbookSheets_(e, source, selectedIds);
}


function installCurrentWorkbookSheets_(e, source, selectedIds) {
  const destinationName = getInstallerDestinationName_(e, source);
  const operationKind = "migrate:" + selectedIds.map(String).sort().join(",");

  return runInstallOperation_(operationKind, source.getId(), destinationName, function() {
    const destinationId = copySendMeBotTemplate_(destinationName);
    const url = "https://docs.google.com/spreadsheets/d/" + destinationId + "/edit";
    try {
      const copied = copySelectedSheets_(source, destinationId, selectedIds);
      markSendMeBotOnboarding_(destinationId, {
        installMode: "migrate",
        sourceSpreadsheetName: source.getName(),
        importedSheets: copied,
        suggestedTrackerSheet: getSuggestedTrackerSheet_(copied)
      });
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
