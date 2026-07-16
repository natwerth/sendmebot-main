"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const source = fs.readFileSync(path.join(ROOT, "Walkthrough.js"), "utf8");
const formSource = fs.readFileSync(path.join(ROOT, "WalkthroughForm.html"), "utf8");
const sandbox = {
  console,
  normalize_: value => String(value || "").trim().toLowerCase(),
  toDisplayHeader_: value => String(value || "").split(" ").map(word =>
    word.charAt(0).toUpperCase() + word.slice(1)
  ).join(" "),
  getTrackerHeaderForTemplate_: value => String(value || "") + " Status",
  getTemplateNameCol_: headers => headers.name || headers["template key"],
  getHeaders_: () => ({ name: 1, subject: 2, body: 3, "attachment link": 4 })
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "Walkthrough.js" });

const templateSheet = {
  getLastRow: () => 4,
  getLastColumn: () => 4,
  getRange: () => ({
    getDisplayValues: () => [
      ["Ready", "Hello {{Name}}", "Hi {{Name}}", ""],
      ["Missing subject", "", "Body", ""],
      ["Missing body", "Subject", "", ""]
    ]
  })
};
assert.deepEqual(
  Array.from(sandbox.getWalkthroughValidTemplateKeys_(templateSheet)),
  ["Ready"]
);

const recordSheet = {
  getLastRow: () => 5,
  getRange: () => ({ getDisplayValues: () => [[""], ["Alice"], [""], ["Bob"]] })
};
assert.equal(
  sandbox.getWalkthroughFirstRecordRow_(recordSheet, { name: 2 }, "Name"),
  3
);
assert.equal(
  sandbox.getWalkthroughFirstRecordRow_(recordSheet, { email: 3 }, "Name"),
  0
);

assert.deepEqual(
  JSON.parse(JSON.stringify(sandbox.buildWalkthroughSampleRowData_(
    { select: 1, name: 2, email: 3, company: 4 },
    "Name",
    "user@example.com"
  ))),
  {
    Select: "",
    Name: "SendMeBot test",
    Email: "user@example.com",
    Company: "Sample Company"
  }
);

assert.deepEqual(
  Array.from(sandbox.getWalkthroughTestVariableHeaders_(
    { select: 1, name: 2, email: 3, status: 4, welcome: 5, company: 6 },
    ["Welcome"],
    "Name"
  )),
  ["Name", "Email", "Company"]
);

assert.match(source, /to:\s*authenticatedEmail/);
assert.match(source, /onboardingState:\s*"complete"/);
assert.match(source, /getAuthenticatedUserEmail_\(\)[\s\S]*MailApp\.sendEmail/);
assert.doesNotMatch(source, /logSentEmail_|stampScheduled|setValue\(true\)/);

const mainScript = formSource.match(/<script>([\s\S]*?)<\/script>/);
assert.ok(mainScript, "walkthrough form script");
assert.doesNotThrow(() => new vm.Script(mainScript[1], { filename: "WalkthroughForm.html" }));
assert.match(formSource, /saveWalkthroughTrackerSetup/);
assert.match(formSource, /sendWalkthroughTestEmail/);
assert.match(formSource, /Finish later/);

console.log("PASS walkthrough enforces valid setup gates and sends a private, non-mutating test");
