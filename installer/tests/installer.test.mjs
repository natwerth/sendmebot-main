import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";

const ROOT = path.resolve(import.meta.dirname, "..");
const sources = ["Config.js", "WorkbookMigration.js", "Installer.js", "Cards.js"]
  .map(name => fs.readFileSync(path.join(ROOT, "src", name), "utf8"));

function sheet(name, id, features = {}) {
  return {
    getName: () => name,
    getSheetId: () => id,
    getProtections: type => type === "RANGE" ? (features.protections || []) : [],
    getCharts: () => features.charts || [],
    getDrawings: () => features.drawings || [],
    getImages: () => features.images || [],
    getSlicers: () => features.slicers || [],
    getDataSourceTables: () => features.dataSources || [],
    getDataSourcePivotTables: () => features.dataPivots || []
  };
}

const sandbox = {
  console,
  Date,
  JSON,
  SpreadsheetApp: {
    ProtectionType: { RANGE: "RANGE", SHEET: "SHEET" },
    getActiveSpreadsheet: () => null,
    openById: id => ({ id })
  },
  Utilities: {
    DigestAlgorithm: { SHA_256: "SHA_256" },
    computeDigest: value => Array.from(Buffer.from(value)),
    base64EncodeWebSafe: value => Buffer.from(value).toString("base64url")
  }
};
vm.createContext(sandbox);
sources.forEach((source, index) => vm.runInContext(source, sandbox, { filename: String(index) }));

const simple = {
  getNamedRanges: () => [],
  getFormUrl: () => "",
  getSheets: () => [sheet("People", 1), sheet("_SendMeBot", 2)]
};
const simpleScan = sandbox.scanWorkbookCompatibility_(simple);
assert.equal(simpleScan.sheets.length, 1);
assert.equal(simpleScan.warnings.length, 0);
assert.match(simpleScan.notices[0], /Macros.*not copied/);

const complex = {
  getNamedRanges: () => [{}, {}],
  getFormUrl: () => "https://docs.google.com/forms/d/form-id",
  getSheets: () => [sheet("People", 1, { charts: [{}], protections: [{}] })]
};
const complexScan = sandbox.scanWorkbookCompatibility_(complex);
assert.equal(complexScan.warnings.some(value => /named range/.test(value)), true);
assert.equal(complexScan.warnings.some(value => /Google Form/.test(value)), true);
assert.equal(complexScan.warnings.some(value => /chart/.test(value)), true);

const destinationNames = new Set(["Tracker", "Tracker (Imported)"]);
const destination = { getSheetByName: name => destinationNames.has(name) ? {} : null };
assert.equal(sandbox.getAvailableImportedSheetName_(destination, "Tracker"), "Tracker (Imported) 2");
assert.equal(sandbox.getAvailableImportedSheetName_(destination, "People"), "People");

assert.equal(sandbox.getSuggestedTrackerSheet_([
  { sourceName: "People", destinationName: "People", hasUsableHeaders: true },
  { sourceName: "Tracker", destinationName: "Tracker (Imported)", hasUsableHeaders: true }
]), "Tracker (Imported)");
assert.equal(sandbox.getSuggestedTrackerSheet_([
  { sourceName: "Notes", destinationName: "Notes", hasUsableHeaders: false },
  { sourceName: "People", destinationName: "People", hasUsableHeaders: true }
]), "People");
assert.equal(sandbox.getSuggestedTrackerSheet_([]), "Tracker");
assert.deepEqual(
  Array.from(sandbox.getAllInstallerSheetIds_({ sheets: [
    { sheetId: 11 }, { sheetId: "22" }, { sheetId: 33 }
  ] })),
  ["11", "22", "33"]
);

let onboardingRows = [];
const onboardingSheet = {
  getLastRow: () => 1,
  getRange: () => ({
    getValues: () => [["appId", "sendmebot"]],
    setValues: values => { onboardingRows = values; }
  }),
  clearContents() {},
  hideSheet() {}
};
sandbox.SpreadsheetApp.openById = () => ({
  getSheetByName: name => name === "_SendMeBot" ? onboardingSheet : null
});
const onboarding = sandbox.markSendMeBotOnboarding_("destination-id", {
  installMode: "migrate",
  sourceSpreadsheetName: "Original",
  importedSheets: [{ destinationName: "People" }],
  suggestedTrackerSheet: "People"
});
assert.equal(onboarding.appId, "sendmebot");
assert.equal(onboarding.onboardingState, "pending");
assert.equal(onboarding.onboardingAutoPrompted, "false");
assert.equal(onboarding.suggestedTrackerSheet, "People");
assert.equal(Object.fromEntries(onboardingRows).sourceSpreadsheetName, "Original");
sandbox.SpreadsheetApp.openById = id => ({ id });

const missingContext = sandbox.getInstallerSheetsContext_({});
assert.equal(missingContext.hasFileScope, false);
assert.equal(missingContext.spreadsheetId, "");
assert.equal(sandbox.getCurrentInstallerSpreadsheet_({}), null);

const deniedContext = sandbox.getInstallerSheetsContext_({
  sheets: { addonHasFileScopePermission: false, id: "must-not-be-used", title: "Private" }
});
assert.equal(deniedContext.hasFileScope, false);
assert.equal(deniedContext.spreadsheetId, "");

const grantedEvent = {
  sheets: {
    addonHasFileScopePermission: true,
    id: "sheet-id",
    title: "Authorized workbook"
  }
};
const grantedContext = sandbox.getInstallerSheetsContext_(grantedEvent);
assert.equal(grantedContext.hasFileScope, true);
assert.equal(grantedContext.spreadsheetId, "sheet-id");
assert.equal(grantedContext.title, "Authorized workbook");
assert.equal(sandbox.getCurrentInstallerSpreadsheet_(grantedEvent).id, "sheet-id");

const cardActionEvent = {
  commonEventObject: {
    parameters: {
      sourceSpreadsheetId: "sheet-id-from-action",
      sourceSpreadsheetName: "Authorized workbook"
    }
  }
};
const cardActionContext = sandbox.getInstallerSheetsContext_(cardActionEvent);
assert.equal(cardActionContext.hasFileScope, true);
assert.equal(cardActionContext.spreadsheetId, "sheet-id-from-action");
assert.equal(cardActionContext.title, "Authorized workbook");
assert.equal(sandbox.getCurrentInstallerSpreadsheet_(cardActionEvent).id, "sheet-id-from-action");

const legacyActionEvent = {
  parameters: {
    sourceSpreadsheetId: "legacy-action-id",
    sourceSpreadsheetName: "Legacy action workbook"
  }
};
assert.equal(sandbox.getCurrentInstallerSpreadsheet_(legacyActionEvent).id, "legacy-action-id");

const activeWorkbook = { getId: () => "active-sheet-id" };
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => activeWorkbook;
sandbox.SpreadsheetApp.openById = () => { throw new Error("openById should not be required"); };
assert.equal(sandbox.getCurrentInstallerSpreadsheet_({
  commonEventObject: {
    parameters: {
      sourceSpreadsheetId: "active-sheet-id",
      sourceSpreadsheetName: "Active workbook"
    }
  }
}), activeWorkbook);
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => null;
sandbox.SpreadsheetApp.openById = id => ({ id });

const fullWorkbook = { getId: () => "full-id", getName: () => "Full workbook" };
let fullInstallIds = [];
sandbox.getCurrentInstallerSpreadsheet_ = () => fullWorkbook;
sandbox.scanWorkbookCompatibility_ = () => ({
  sheets: [{ sheetId: 1 }, { sheetId: 2 }, { sheetId: 3 }],
  warnings: [],
  notices: []
});
sandbox.installCurrentWorkbookSheets_ = (event, source, ids) => {
  assert.equal(source, fullWorkbook);
  fullInstallIds = Array.from(ids);
  return "installed";
};
assert.equal(sandbox.installEntireCurrentWorkbook({}), "installed");
assert.deepEqual(fullInstallIds, ["1", "2", "3"]);

sandbox.scanWorkbookCompatibility_ = () => ({
  sheets: [{ sheetId: 1 }],
  warnings: ["Review this feature."],
  notices: []
});
sandbox.buildEntireWorkbookWarningResponse_ = () => "review-card";
assert.equal(sandbox.installEntireCurrentWorkbook({}), "review-card");

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, "src", "appsscript.json"), "utf8"));
assert.ok(manifest.oauthScopes.includes("https://www.googleapis.com/auth/script.projects"));
assert.ok(manifest.oauthScopes.includes("https://www.googleapis.com/auth/drive.file"));
assert.equal(JSON.stringify(manifest).includes("Gmail"), false);

const allSource = sources.join("\n");
assert.doesNotMatch(allSource, /projects\.create|SENDMEBOT_SOURCE_SCRIPT_ID/);
assert.match(allSource, /Install SendMeBot into /);
assert.match(allSource, /sourceName \+ " — SendMeBot"/);
assert.match(allSource, /destinationWorkbookName/);
assert.match(allSource, /markSendMeBotOnboarding_/);
assert.match(allSource, /setParameters\(migrationParameters\)/);
assert.match(allSource, /commonEventObject && event\.commonEventObject\.parameters/);
assert.match(allSource, /setFunctionName\("installEntireCurrentWorkbook"\)/);
assert.match(allSource, /Choose which sheets to copy/);
assert.match(allSource, /copy all worksheet tabs/i);
assert.match(allSource, /buildEntireWorkbookWarningResponse_/);
assert.match(allSource, /original spreadsheet will not be changed/i);
assert.match(allSource, /getActiveSpreadsheet\s*\(\)/);
assert.match(allSource, /active\.getId\(\)[\s\S]*context\.spreadsheetId/);
assert.match(allSource, /addonHasFileScopePermission\s*===\s*true/);

console.log("PASS installer uses explicit file scope, scans safely, handles collisions, and contains no legacy project injection");
