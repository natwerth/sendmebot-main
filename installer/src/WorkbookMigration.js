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


function copySelectedSheets_(source, destinationId, selectedIds) {
  const destination = SpreadsheetApp.openById(destinationId);
  const selected = {};
  (selectedIds || []).forEach(id => { selected[String(id)] = true; });
  const copied = [];

  source.getSheets().forEach(sheet => {
    if (!selected[String(sheet.getSheetId())]) return;
    const response = Sheets.Spreadsheets.Sheets.copyTo(
      { destinationSpreadsheetId: destinationId },
      source.getId(),
      sheet.getSheetId()
    );
    const copiedSheet = destination.getSheetById(response.sheetId);
    const finalName = getAvailableImportedSheetName_(destination, sheet.getName());
    copiedSheet.setName(finalName);
    copied.push({ sourceName: sheet.getName(), destinationName: finalName, sheetId: response.sheetId });
  });
  return copied;
}


function auditCopiedSheets_(destinationId, copied) {
  const destination = SpreadsheetApp.openById(destinationId);
  const warnings = [];
  (copied || []).forEach(item => {
    const sheet = destination.getSheetByName(item.destinationName);
    if (!sheet) {
      warnings.push(item.destinationName + " could not be verified after copying.");
      return;
    }
    const range = sheet.getDataRange();
    const formulas = range.getFormulas();
    const displays = range.getDisplayValues();
    let broken = 0;
    formulas.forEach((row, rowIndex) => row.forEach((formula, colIndex) => {
      if (/#REF!/i.test(String(formula || "")) || /#REF!/i.test(String(displays[rowIndex][colIndex] || ""))) {
        broken++;
      }
    }));
    if (broken) warnings.push(item.destinationName + " contains " + broken + " broken reference(s).");
  });
  return warnings;
}
