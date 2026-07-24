function safelyCountWorkbookFeature_(callback) {
  try {
    const result = callback();
    return Array.isArray(result) ? result.length : Number(result || 0);
  } catch (err) {
    return 0;
  }
}


function scanWorkbookCompatibility_(ss) {
  const warnings = [];
  const notices = [];
  const namedRangeCount = safelyCountWorkbookFeature_(() => ss.getNamedRanges());
  if (namedRangeCount) warnings.push(namedRangeCount + " named range(s) require manual verification.");
  if (safelyCountWorkbookFeature_(() => ss.getFormUrl() ? 1 : 0)) {
    warnings.push("The source is linked to a Google Form; that form link will not be transferred.");
  }

  const sheets = ss.getSheets()
    .filter(sheet => sheet.getName() !== "_SendMeBot")
    .map(sheet => {
      const features = [];
      const featureChecks = [
        ["protected range(s)", () => sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)],
        ["protected sheet setting(s)", () => sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)],
        ["chart(s)", () => sheet.getCharts()],
        ["drawing(s)", () => sheet.getDrawings()],
        ["over-grid image(s)", () => sheet.getImages()],
        ["slicer(s)", () => sheet.getSlicers()],
        ["data source table(s)", () => sheet.getDataSourceTables()],
        ["data source pivot table(s)", () => sheet.getDataSourcePivotTables()]
      ];
      featureChecks.forEach(([label, callback]) => {
        const count = safelyCountWorkbookFeature_(callback);
        if (count) features.push(count + " " + label);
      });
      if (features.length) warnings.push(sheet.getName() + ": " + features.join(", ") + ".");
      return {
        name: sheet.getName(),
        sheetId: sheet.getSheetId(),
        features: features
      };
    });

  notices.push(
    "Macros and any script bound to the source spreadsheet are not copied; SendMeBot supplies its own bound script."
  );
  return { sheets: sheets, warnings: warnings, notices: notices };
}


function getAvailableImportedSheetName_(destination, requestedName) {
  const rawBase = SENDMEBOT_RESERVED_SHEET_NAMES.indexOf(requestedName) !== -1
    ? requestedName + " (Imported)"
    : requestedName;
  const base = rawBase.slice(0, 100);
  let candidate = base;
  let index = 2;
  while (destination.getSheetByName(candidate)) {
    const suffix = " " + index;
    candidate = base.slice(0, 100 - suffix.length) + suffix;
    index++;
  }
  return candidate;
}


function getAvailableImportedSheetNameFromMap_(existingNames, requestedName) {
  const rawBase = SENDMEBOT_RESERVED_SHEET_NAMES.indexOf(requestedName) !== -1
    ? requestedName + " (Imported)"
    : requestedName;
  const base = rawBase.slice(0, 100);
  let candidate = base;
  let index = 2;
  while (existingNames[candidate]) {
    const suffix = " " + index;
    candidate = base.slice(0, 100 - suffix.length) + suffix;
    index++;
  }
  return candidate;
}


function getInstallerSpreadsheetSheetsByApi_(spreadsheetId) {
  const response = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets.properties(sheetId,title,hidden,gridProperties)"
  });
  return (response.sheets || []).map(sheet => sheet.properties || {});
}


function getInstallerColumnLetter_(columnNumber) {
  let value = Math.max(Number(columnNumber) || 1, 1);
  let result = "";
  while (value > 0) {
    value--;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}


function renameInstallerSheetByApi_(spreadsheetId, sheetId, title) {
  Sheets.Spreadsheets.batchUpdate({
    requests: [{
      updateSheetProperties: {
        properties: { sheetId: sheetId, title: title },
        fields: "title"
      }
    }]
  }, spreadsheetId);
}


function copySelectedSheets_(source, destinationId, selectedIds) {
  const selected = {};
  (selectedIds || []).forEach(id => { selected[String(id)] = true; });
  const copied = [];
  const warnings = [];
  const destinationNames = {};
  getInstallerSpreadsheetSheetsByApi_(destinationId).forEach(properties => {
    destinationNames[properties.title] = true;
  });

  source.getSheets().forEach(sheet => {
    if (!selected[String(sheet.getSheetId())]) return;
    try {
      const response = Sheets.Spreadsheets.Sheets.copyTo(
        { destinationSpreadsheetId: destinationId },
        source.getId(),
        sheet.getSheetId()
      );
      const finalName = getAvailableImportedSheetNameFromMap_(destinationNames, sheet.getName());
      renameInstallerSheetByApi_(destinationId, response.sheetId, finalName);
      destinationNames[finalName] = true;
      const headers = sheet.getLastColumn() > 0
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
        : [];
      copied.push({
        sourceName: sheet.getName(),
        destinationName: finalName,
        sheetId: response.sheetId,
        hasUsableHeaders: headers.some(value => normalizeInstallerValue_(value))
      });
    } catch (err) {
      warnings.push(sheet.getName() + " could not be copied: " + (err.message || err));
    }
  });
  return { copied: copied, warnings: warnings };
}


function getSuggestedTrackerSheet_(copied) {
  const sheets = Array.isArray(copied) ? copied : [];
  const namedTracker = sheets.find(item =>
    normalizeInstallerValue_(item.sourceName).toLowerCase() === "tracker"
  );
  if (namedTracker) return namedTracker.destinationName;
  const usable = sheets.find(item => item.hasUsableHeaders);
  return usable ? usable.destinationName : "Tracker";
}


function markSendMeBotOnboarding_(destinationId, metadata) {
  const registryName = "_SendMeBot";
  const registryRange = "'_SendMeBot'!A:B";
  let registrySheet = getInstallerSpreadsheetSheetsByApi_(destinationId)
    .find(properties => properties.title === registryName);

  if (!registrySheet) {
    const response = Sheets.Spreadsheets.batchUpdate({
      requests: [{ addSheet: { properties: { title: registryName, hidden: true } } }]
    }, destinationId);
    registrySheet = response.replies[0].addSheet.properties;
  } else if (!registrySheet.hidden) {
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        updateSheetProperties: {
          properties: { sheetId: registrySheet.sheetId, hidden: true },
          fields: "hidden"
        }
      }]
    }, destinationId);
  }

  const existing = {};
  const registryValues = Sheets.Spreadsheets.Values.get(destinationId, registryRange).values || [];
  registryValues.forEach(row => {
    const key = normalizeInstallerValue_(row[0]);
    if (key) existing[key] = String(row[1] || "");
  });

  const settings = metadata || {};
  const values = Object.assign({}, existing, {
    onboardingState: "pending",
    onboardingAutoPrompted: "false",
    installMode: normalizeInstallerValue_(settings.installMode) || "new",
    sourceSpreadsheetName: normalizeInstallerValue_(settings.sourceSpreadsheetName),
    importedSheets: JSON.stringify(settings.importedSheets || []),
    suggestedTrackerSheet: normalizeInstallerValue_(settings.suggestedTrackerSheet) || "Tracker",
    onboardingCreatedAt: new Date().toISOString(),
    onboardingCompletedAt: "",
    testEmailSentAt: ""
  });
  const rows = Object.keys(values).sort().map(key => [key, values[key]]);
  Sheets.Spreadsheets.Values.clear({}, destinationId, registryRange);
  Sheets.Spreadsheets.Values.update(
    { values: rows },
    destinationId,
    "'_SendMeBot'!A1:B" + rows.length,
    { valueInputOption: "RAW" }
  );
  return values;
}


function auditCopiedSheets_(destinationId, copied) {
  const warnings = [];
  const destinationSheets = {};
  getInstallerSpreadsheetSheetsByApi_(destinationId).forEach(properties => {
    destinationSheets[properties.title] = properties;
  });

  (copied || []).forEach(item => {
    const sheet = destinationSheets[item.destinationName];
    if (!sheet) {
      warnings.push(item.destinationName + " could not be verified after copying.");
      return;
    }
    const escapedName = String(item.destinationName || "").replace(/'/g, "''");
    const columnCount = sheet.gridProperties && sheet.gridProperties.columnCount;
    const range = "'" + escapedName + "'!A:" + getInstallerColumnLetter_(columnCount || 26);
    const formulas = Sheets.Spreadsheets.Values.get(destinationId, range, {
      valueRenderOption: "FORMULA"
    }).values || [];
    const displays = Sheets.Spreadsheets.Values.get(destinationId, range, {
      valueRenderOption: "FORMATTED_VALUE"
    }).values || [];
    let broken = 0;
    formulas.forEach((row, rowIndex) => row.forEach((formula, colIndex) => {
      const display = displays[rowIndex] && displays[rowIndex][colIndex];
      if (/#REF!/i.test(String(formula || "")) || /#REF!/i.test(String(display || ""))) {
        broken++;
      }
    }));
    if (broken) warnings.push(item.destinationName + " contains " + broken + " broken reference(s).");
  });
  return warnings;
}
