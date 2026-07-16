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

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, "src", "appsscript.json"), "utf8"));
assert.ok(manifest.oauthScopes.includes("https://www.googleapis.com/auth/script.projects"));
assert.ok(manifest.oauthScopes.includes("https://www.googleapis.com/auth/drive.file"));
assert.equal(JSON.stringify(manifest).includes("Gmail"), false);

const allSource = sources.join("\n");
assert.doesNotMatch(allSource, /projects\.create|SENDMEBOT_SOURCE_SCRIPT_ID/);
assert.match(allSource, /Install SendMeBot into /);
assert.match(allSource, /original spreadsheet will not be changed/i);
assert.doesNotMatch(allSource, /getActiveSpreadsheet\s*\(/);
assert.match(allSource, /addonHasFileScopePermission\s*===\s*true/);

console.log("PASS installer uses explicit file scope, scans safely, handles collisions, and contains no legacy project injection");
