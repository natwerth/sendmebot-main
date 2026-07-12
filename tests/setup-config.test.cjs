"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const codeSource = fs.readFileSync(path.join(ROOT, "Code.js"), "utf8");
const setupSource = fs.readFileSync(path.join(ROOT, "SetupForm.html"), "utf8");
const modalSource = fs.readFileSync(path.join(ROOT, "Modal.html"), "utf8");
const modalFiles = ["AddImage.html", "AddSender.html", "SendForm.html", "SetupForm.html", "TemplateForm.html"];
const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

function makeEnvironment(initialProperties) {
  const values = Object.assign({}, initialProperties || {});
  const properties = {
    getProperty(key) { return values[key] || null; },
    setProperties(updates) { Object.assign(values, updates); }
  };
  const tracker = {
    getName() { return "Tracker"; },
    getLastColumn() { return 3; },
    getRange() {
      return {
        getDisplayValues() { return [["Select", "Student Name", "Email"]]; },
        getValues() { return [["Select", "Student Name", "Email"]]; }
      };
    }
  };
  const spreadsheet = {
    getSheetByName(name) { return name === "Tracker" ? tracker : null; },
    getSheets() { return [tracker]; }
  };
  const lock = { waitLock() {}, releaseLock() {} };
  const sandbox = {
    console,
    PropertiesService: { getDocumentProperties: () => properties },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    LockService: { getDocumentLock: () => lock },
    normalize_: value => String(value || "").trim().toLowerCase(),
    getHeaders_() {
      return { select: 1, "student name": 2, email: 3 };
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(codeSource, sandbox, { filename: "Code.js" });
  sandbox.normalize_ = value => String(value || "").trim().toLowerCase();
  sandbox.getHeaders_ = () => ({ select: 1, "student name": 2, email: 3 });
  return { sandbox, values, spreadsheet };
}

test("legacy configuration defaults remain available", () => {
  const { sandbox } = makeEnvironment();
  const config = sandbox.getSendMeBotConfig_();
  assert.equal(config.trackerSheetName, "Hires & Conversion - Intern");
  assert.equal(config.recordIdHeader, "Student Name");
});

test("Setup saves shared tracker and record-ID choices", () => {
  const { sandbox, values } = makeEnvironment();
  const result = sandbox.saveSendMeBotSetup({
    trackerSheetName: "Tracker",
    recordIdHeader: "Student Name"
  });
  assert.equal(result.trackerSheetName, "Tracker");
  assert.equal(values.SENDMEBOT_TRACKER_SHEET, "Tracker");
  assert.equal(values.SENDMEBOT_RECORD_ID_HEADER, "Student Name");
  assert.equal(sandbox.getTrackerSheet_().getName(), "Tracker");
});

test("Setup does not require a general Status column", () => {
  const { sandbox } = makeEnvironment();
  assert.doesNotThrow(() => sandbox.saveSendMeBotSetup({
    trackerSheetName: "Tracker",
    recordIdHeader: "Student Name"
  }));
  assert.doesNotMatch(codeSource, /missing the \"Status\" column/);
});

test("Setup rejects missing required Tracker columns", () => {
  const { sandbox } = makeEnvironment();
  sandbox.getHeaders_ = () => ({ "student name": 2, status: 4 });
  assert.throws(() => sandbox.saveSendMeBotSetup({
    trackerSheetName: "Tracker",
    recordIdHeader: "Student Name"
  }), /Select/);
});

test("Setup dialog exposes sheet and record-ID selectors without named ranges", () => {
  assert.match(setupSource, /id="trackerSheet"/);
  assert.match(setupSource, /id="recordIdHeader"/);
  assert.match(codeSource, /\.addItem\("Setup", "openSetupForm"\)/);
  assert.doesNotMatch(codeSource + setupSource, /setNamedRange|getRangeByName/);
  assert.match(setupSource, /includeHtml_\("Modal"\)/);
  assert.match(modalSource, /window\.SendMeBotModal/);
  const script = setupSource.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace("const context = <?!= contextJson ?>;", "const context = {config:{},sheets:[]};")
    .replace(/\s*renderHeaders\(\);\s*$/, "");
  const element = () => ({ value: "", innerHTML: "", appendChild() {}, textContent: "" });
  const sandbox = {
    console,
    document: { getElementById: element, createElement: element },
    google: { script: { host: {}, run: {} } }
  };
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox, { filename: "SetupForm.html" });
});

test("modal dialogs use the shared include and its utilities are syntactically valid", () => {
  modalFiles.forEach(file => {
    const source = fs.readFileSync(path.join(ROOT, file), "utf8");
    assert.match(source, /includeHtml_\("Modal"\)/, file);
    const mainScript = source.match(/<script>([\s\S]*?)<\/script>/);
    assert.ok(mainScript, file + " main script");
    assert.doesNotThrow(() => new vm.Script(
      mainScript[1].replace(/<\?!=[\s\S]*?\?>/g, "null"),
      { filename: file }
    ));
  });
  assert.match(codeSource, /createTemplateFromFile\("AddImage"\)[\s\S]*?\.evaluate\(\)/);
  assert.match(codeSource, /function includeHtml_\(filename\)/);
  const script = modalSource.match(/<script>([\s\S]*?)<\/script>/)[1];
  const sandbox = { window: {}, document: { getElementById() { return null; } } };
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox, { filename: "Modal.html" });
  assert.equal(sandbox.window.SendMeBotModal.errorText({}, "Fallback"), "Fallback");
  assert.doesNotMatch(
    fs.readFileSync(path.join(ROOT, "AssistantSidebar.html"), "utf8"),
    /includeHtml_\("Modal"\)/
  );
});

let failures = 0;
tests.forEach(({ name, fn }) => {
  try {
    fn();
    console.log("PASS " + name);
  } catch (err) {
    failures++;
    console.error("FAIL " + name);
    console.error(err && err.stack ? err.stack : err);
  }
});
console.log("\n" + (tests.length - failures) + "/" + tests.length + " setup tests passed.");
if (failures) process.exitCode = 1;
