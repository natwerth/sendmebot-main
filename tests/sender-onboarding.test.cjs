"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const dataSource = fs.readFileSync(path.join(ROOT, "Data.js"), "utf8");
const emailsSource = fs.readFileSync(path.join(ROOT, "Emails.js"), "utf8");
const formSource = fs.readFileSync(path.join(ROOT, "SendForm.html"), "utf8");
const addSenderSource = fs.readFileSync(path.join(ROOT, "AddSender.html"), "utf8");

function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

function createLock() {
  return {
    waits: 0,
    releases: 0,
    waitLock() { this.waits++; },
    releaseLock() { this.releases++; },
    tryLock() { return true; }
  };
}

function createSheet(headers, rows) {
  return {
    headers: headers.slice(),
    rows: rows.map(row => row.slice()),
    writes: [],
    getLastRow() { return this.rows.length + 1; },
    getLastColumn() { return this.headers.length; },
    getName() { return "Senders"; },
    getRange(row, col, numRows, numCols) {
      const sheet = this;
      const height = numRows || 1;
      const width = numCols || 1;
      function valueAt(r, c) {
        return r === 1 ? sheet.headers[c - 1] : (sheet.rows[r - 2] || [])[c - 1];
      }
      function setAt(r, c, value) {
        if (r === 1) {
          sheet.headers[c - 1] = value;
        } else {
          while (sheet.rows.length < r - 1) sheet.rows.push(sheet.headers.map(() => ""));
          sheet.rows[r - 2][c - 1] = value;
        }
        sheet.writes.push({ row: r, col: c, value });
      }
      return {
        getValues() {
          return Array.from({ length: height }, (_, r) =>
            Array.from({ length: width }, (_, c) => valueAt(row + r, col + c))
          );
        },
        getDisplayValue() { return String(valueAt(row, col) || ""); },
        getRichTextValue() {
          const value = String(valueAt(row, col) || "");
          return { getText() { return value; } };
        },
        setValue(value) { setAt(row, col, value); return this; },
        setValues(values) {
          values.forEach((valuesRow, r) => valuesRow.forEach((value, c) => setAt(row + r, col + c, value)));
          return this;
        },
        setRichTextValue(value) {
          setAt(row, col, value && value.text !== undefined ? value.text : String(value || ""));
          return this;
        },
        setWrap() { return this; }
      };
    }
  };
}

function createDataEnvironment(senderRows, authenticatedEmail = "owner@akamai.com") {
  const senderSheet = createSheet(
    ["Name", "Email", "Signature", "", "Name", "Link", "Width"],
    senderRows || []
  );
  const lock = createLock();
  const spreadsheet = {
    toasts: [],
    getSheetByName(name) { return name === "Senders" ? senderSheet : null; },
    insertSheet() { return senderSheet; },
    toast(message) { this.toasts.push(message); }
  };
  const sandbox = {
    console,
    Session: {
      getEffectiveUser() { return { getEmail: () => authenticatedEmail }; },
      getScriptTimeZone() { return "America/Chicago"; }
    },
    SpreadsheetApp: { getActiveSpreadsheet() { return spreadsheet; } },
    LockService: { getDocumentLock() { return lock; } },
    normalize_: normalize,
    buildRichTextValueFromTemplateHtml_(value) { return { text: value }; },
    Logger: { log() {} }
  };
  vm.createContext(sandbox);
  vm.runInContext(dataSource, sandbox, { filename: "Data.js" });
  sandbox.normalize_ = normalize;
  sandbox.buildRichTextValueFromTemplateHtml_ = value => ({ text: value });
  return { sandbox, senderSheet, spreadsheet, lock };
}

function createTriggerApp(initialTriggers, options) {
  const config = options || {};
  const triggers = initialTriggers ? initialTriggers.slice() : [];
  const app = {
    EventType: { CLOCK: "CLOCK" },
    createAttempts: 0,
    getProjectTriggers() { return triggers.slice(); },
    newTrigger(handler) {
      this.createAttempts++;
      if (config.failCreate) throw new Error("authorization required");
      return {
        timeBased() { return this; },
        everyMinutes() { return this; },
        create() {
          triggers.push({
            getHandlerFunction: () => handler,
            getEventType: () => "CLOCK"
          });
          return triggers[triggers.length - 1];
        }
      };
    }
  };
  return app;
}

function trigger(handler = "sendScheduledEmails", eventType = "CLOCK") {
  return {
    getHandlerFunction: () => handler,
    getEventType: () => eventType
  };
}

function loadEmailsSandbox(extra) {
  const sandbox = Object.assign({
    console,
    Date,
    JSON,
    Logger: { log() {} },
    Session: {
      getEffectiveUser() { return { getEmail: () => "owner@akamai.com" }; },
      getScriptTimeZone() { return "America/Chicago"; }
    },
    LockService: {
      getUserLock() { return createLock(); },
      getDocumentLock() { return createLock(); },
      getScriptLock() { return createLock(); }
    },
    SpreadsheetApp: { getActiveSpreadsheet() { return {}; }, flush() {} },
    ScriptApp: createTriggerApp([]),
    MailApp: { sendEmail() { throw new Error("Unexpected real-mail path"); } },
    normalize_: normalize
  }, extra || {});
  vm.createContext(sandbox);
  vm.runInContext(emailsSource, sandbox, { filename: "Emails.js" });
  sandbox.normalize_ = normalize;
  return sandbox;
}

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

test("authenticated sender state returns only the matching record", () => {
  const { sandbox, spreadsheet } = createDataEnvironment([
    ["Owner Name", "owner@akamai.com", "Owner signature", "", "Logo", "file1", "320"],
    ["Other User", "other@akamai.com", "Other signature", "", "Banner", "file2", "400"]
  ]);
  const state = sandbox.getCurrentUserSenderState_(spreadsheet);
  assert.equal(state.status, "ready");
  assert.equal(state.record.name, "Owner Name");
  assert.equal(state.record.email, "owner@akamai.com");
  assert.equal(JSON.stringify(state).includes("other@akamai.com"), false);
});

test("missing and duplicate current-user sender states are explicit", () => {
  const missing = createDataEnvironment([
    ["Other", "other@akamai.com", "", "", "Logo", "file1", "320"]
  ]);
  assert.equal(missing.sandbox.getCurrentUserSenderState_(missing.spreadsheet).status, "missing");

  const duplicate = createDataEnvironment([
    ["Owner A", "owner@akamai.com", "A", "", "Logo", "file1", "320"],
    ["Owner B", "owner@akamai.com", "B", "", "Banner", "file2", "400"]
  ]);
  const state = duplicate.sandbox.getCurrentUserSenderState_(duplicate.spreadsheet);
  assert.equal(state.status, "duplicate");
  assert.equal(state.duplicateCount, 2);
  assert.equal(state.record, null);
});

test("sender onboarding derives email server-side and writes only A:C", () => {
  const { sandbox, senderSheet, spreadsheet, lock } = createDataEnvironment([
    ["", "", "", "separator", "Logo", "drive-link", "320"]
  ]);
  const imageBefore = senderSheet.rows.map(row => row.slice(3, 7));
  const result = sandbox.saveSenderProfile_({
    name: "Owner Name",
    signature: "All the best"
  }, {
    spreadsheet,
    authenticatedEmail: "owner@akamai.com",
    lock,
    buildRichText: value => ({ text: value }),
    suppressToast: true
  });
  assert.equal(result.senderState.record.email, "owner@akamai.com");
  assert.deepEqual(senderSheet.rows[0].slice(0, 3), ["Owner Name", "owner@akamai.com", "All the best"]);
  assert.deepEqual(senderSheet.rows.map(row => row.slice(3, 7)), imageBefore);
  assert.equal(senderSheet.writes.some(write => write.col >= 4), false);
});

test("client-submitted alternate sender email is rejected", () => {
  const { sandbox, spreadsheet, lock } = createDataEnvironment([]);
  assert.throws(() => sandbox.saveSenderProfile_({
    name: "Owner",
    email: "other@akamai.com",
    signature: ""
  }, {
    spreadsheet,
    authenticatedEmail: "owner@akamai.com",
    lock,
    suppressToast: true
  }), /authenticated Google account/);

  assert.throws(() => sandbox.saveSenderProfile({
    name: "Owner",
    email: "other@akamai.com",
    signature: ""
  }, {
    authenticatedEmail: "other@akamai.com",
    spreadsheet,
    lock
  }), /authenticated Google account/);
});

test("SendForm bootstrap exposes current sender but not another user's row", () => {
  const data = createDataEnvironment([
    ["Owner Name", "owner@akamai.com", "Owner signature", "", "Logo", "file1", "320"],
    ["Other", "other@akamai.com", "Other signature", "", "", "", ""]
  ]);
  const tracker = {};
  const templateSheet = {};
  data.spreadsheet.getSheetByName = name => {
    if (name === "Senders") return data.senderSheet;
    if (name === "Templates") return templateSheet;
    return null;
  };
  const sandbox = Object.assign(data.sandbox, {
    ScriptApp: createTriggerApp([trigger()]),
    getTrackerSheet_: () => tracker,
    getHeaders_: () => ({ select: 1 }),
    getSelectedRows_: () => [2, 3, 4, 5],
    getTemplateKeys_: () => ["Welcome"],
    getRecipientFieldsForSendForm_: () => [{ header: "Email", label: "Email" }]
  });
  vm.runInContext(emailsSource, sandbox, { filename: "Emails.js" });
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getHeaders_ = () => ({ select: 1 });
  sandbox.getSelectedRows_ = () => [2, 3, 4, 5];
  sandbox.getTemplateKeys_ = () => ["Welcome"];
  sandbox.getRecipientFieldsForSendForm_ = () => [{ header: "Email", label: "Email" }];
  const context = sandbox.getSendFormOpenContext_("schedule");
  assert.equal(context.authenticatedEmail, "owner@akamai.com");
  assert.equal(context.senderState.record.name, "Owner Name");
  assert.equal(context.selectedRowCount, 4);
  assert.equal(JSON.stringify(context).includes("other@akamai.com"), false);
});

test("existing trigger is reused and missing trigger is created exactly once", () => {
  const sandbox = loadEmailsSandbox();
  const readyApp = createTriggerApp([trigger()]);
  const readyState = sandbox.ensureCurrentUserScheduledTrigger_({
    scriptApp: readyApp,
    lock: createLock()
  });
  assert.equal(readyState.status, "ready");
  assert.equal(readyApp.createAttempts, 0);

  const missingApp = createTriggerApp([]);
  const setupLock = createLock();
  sandbox.ensureCurrentUserScheduledTrigger_({ scriptApp: missingApp, lock: setupLock });
  sandbox.ensureCurrentUserScheduledTrigger_({ scriptApp: missingApp, lock: setupLock });
  assert.equal(missingApp.createAttempts, 1);
  assert.equal(missingApp.getProjectTriggers().length, 1);
  assert.equal(setupLock.waits, 2);
  assert.equal(setupLock.releases, 2);
});

test("duplicate or failed trigger creation blocks safely", () => {
  const sandbox = loadEmailsSandbox();
  assert.throws(() => sandbox.ensureCurrentUserScheduledTrigger_({
    scriptApp: createTriggerApp([trigger(), trigger()]),
    lock: createLock()
  }), /Multiple sendScheduledEmails triggers/);
  assert.throws(() => sandbox.ensureCurrentUserScheduledTrigger_({
    scriptApp: createTriggerApp([], { failCreate: true }),
    lock: createLock()
  }), /could not be authorized or enabled/);
});

test("trigger setup failure occurs before queue, tracker, or checkbox writes", () => {
  const tracker = {};
  const spreadsheet = { getSheetByName() { return {}; } };
  const sandbox = loadEmailsSandbox({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, flush() {} }
  });
  const calls = { tracker: 0, queue: 0, checkbox: 0 };
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getHeaders_ = () => ({ select: 1 });
  sandbox.getSelectedRows_ = () => [2];
  sandbox.requireAuthenticatedSenderProfile_ = () => ({ email: "owner@akamai.com" });
  const prepared = { 2: { subject: "Prepared once", metadata: {} } };
  sandbox.preflightScheduledRows_ = () => ({
    validRows: [2], failures: [], preparedRows: prepared
  });
  sandbox.ensureCurrentUserScheduledTrigger_ = () => { throw new Error("authorization required"); };
  sandbox.ensureTrackerColumnForTemplate_ = () => { calls.tracker++; };
  sandbox.saveQueuedJob_ = () => { calls.queue++; };
  sandbox.clearSelectedRow_ = () => { calls.checkbox++; };

  assert.throws(() => sandbox.queueSendFormJob({
    action: "schedule",
    template: "Welcome",
    toField: { value: "Email", valueType: "field" },
    scheduledForIso: "2099-01-01T14:00:00.000Z",
    scheduledTimeZone: "America/Chicago",
    scheduledDisplayText: "Jan 1, 2099, 8:00 AM CST"
  }), /authorization required/);
  assert.deepEqual(calls, { tracker: 0, queue: 0, checkbox: 0 });
  const preflightSource = emailsSource.slice(
    emailsSource.indexOf("function preflightScheduledRows_"),
    emailsSource.indexOf("function queueSendFormJob")
  );
  assert.doesNotMatch(preflightSource, /getTemplateStatusColumn_/);
});

test("document-lock failure leaves selections and tracker statuses intact", () => {
  const sandbox = loadEmailsSandbox({
    LockService: {
      getDocumentLock() { return { tryLock() { return false; }, releaseLock() {} }; },
      getUserLock() { return createLock(); },
      getScriptLock() { return createLock(); }
    }
  });
  let cleared = 0;
  let marked = 0;
  sandbox.clearSelectedRow_ = () => { cleared++; };
  sandbox.markSchedulingRowsStarted_ = () => { marked++; };
  const result = sandbox.processQueuedJobs_({});
  assert.equal(result.locked, true);
  assert.equal(cleared, 0);
  assert.equal(marked, 0);
});

test("Scheduling status is batch-written and flushed before row processing", () => {
  const events = [];
  const tracker = {
    getRange(row, col) {
      return { getA1Notation: () => "R" + row + "C" + col };
    },
    getRangeList(ranges) {
      return {
        setValue(value) { events.push({ type: "batch", ranges, value }); }
      };
    }
  };
  const spreadsheet = {
    getSheetByName(name) { return name === "Templates" ? {} : null; },
    toast() {}
  };
  const sandbox = loadEmailsSandbox({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      flush() { events.push({ type: "flush" }); }
    }
  });
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getHeaders_ = () => ({ status: 1, select: 3 });
  sandbox.getTemplateStatusColumn_ = () => 2;
  sandbox.clearSelectedRow_ = () => {};
  sandbox.scheduleOneRow_ = (ss, sheet, row) => {
    events.push({ type: "row", row });
    return { status: "scheduled", row, error: "" };
  };

  sandbox.processOneQueuedJob_({
    action: "schedule", rows: [2, 5], attemptedRows: [2, 5],
    template: "Welcome", sender: "owner@akamai.com",
    scheduledForIso: "2099-01-01T14:00:00.000Z"
  }, {});

  assert.deepEqual(events.slice(0, 4), [
    { type: "batch", ranges: ["R2C2", "R5C2"], value: "Scheduling..." },
    { type: "batch", ranges: ["R2C1", "R5C1"], value: "Scheduling..." },
    { type: "flush" },
    { type: "row", row: 2 }
  ]);
});

test("scheduling writes its job only after trigger readiness is verified", () => {
  const tracker = {};
  const spreadsheet = { getSheetByName() { return {}; } };
  const sandbox = loadEmailsSandbox({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, flush() {} }
  });
  let triggerReady = false;
  let queued = false;
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getHeaders_ = () => ({ select: 1 });
  sandbox.getSelectedRows_ = () => [2];
  sandbox.requireAuthenticatedSenderProfile_ = () => ({ email: "owner@akamai.com" });
  const prepared = { 2: { subject: "Prepared once", metadata: {} } };
  sandbox.preflightScheduledRows_ = () => ({
    validRows: [2], failures: [], preparedRows: prepared
  });
  sandbox.ensureCurrentUserScheduledTrigger_ = () => {
    triggerReady = true;
    return { status: "ready" };
  };
  sandbox.ensureTrackerColumnForTemplate_ = () => {
    assert.equal(triggerReady, true);
    return 2;
  };
  sandbox.getTemplateStatusColumn_ = () => 2;
  sandbox.saveQueuedJob_ = () => {
    assert.equal(triggerReady, true);
    queued = true;
  };
  sandbox.processQueuedJobs_ = preparedJobs => {
    assert.equal(Object.values(preparedJobs)[0], prepared);
    return {
      attempted: 1, successful: 1, sent: 0, scheduled: 1, failed: 0,
      successfulRows: [2], failedRows: [], errors: [], skipped: 0, locked: false
    };
  };

  const result = sandbox.queueSendFormJob({
    action: "schedule",
    template: "Welcome",
    toField: { value: "Email", valueType: "field" },
    scheduledForIso: "2099-01-01T14:00:00.000Z",
    scheduledTimeZone: "America/Chicago",
    scheduledDisplayText: "Jan 1, 2099, 8:00 AM CST"
  });
  assert.equal(queued, true);
  assert.equal(result.scheduled, 1);
});

test("row-level immediate and scheduled attempts all clear their selections", () => {
  const tracker = {};
  const spreadsheet = {
    getSheetByName(name) { return name === "Templates" ? {} : null; },
    toast() {}
  };
  const sandbox = loadEmailsSandbox({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, flush() {} }
  });
  const cleared = [];
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getHeaders_ = () => ({ select: 1 });
  sandbox.clearSelectedRow_ = (sheet, headers, row) => cleared.push(row);
  sandbox.sendOneRowNow_ = (ss, sheet, row) => ({
    status: row === 2 ? "sent" : "failed",
    row,
    error: row === 2 ? "" : "fake failure"
  });

  const immediate = sandbox.processOneQueuedJob_({
    action: "send_now", rows: [2, 3, 6], attemptedRows: [2, 3, 6],
    template: "Welcome", sender: "owner@akamai.com"
  });
  assert.deepEqual(cleared, [2, 3, 6]);
  assert.equal(immediate.successful, 1);
  assert.equal(immediate.failed, 2);
  assert.deepEqual(Array.from(immediate.successfulRows), [2]);
  assert.deepEqual(Array.from(immediate.failedRows), [3, 6]);

  cleared.length = 0;
  sandbox.scheduleOneRow_ = (ss, sheet, row) => ({
    status: row === 4 ? "scheduled" : "failed",
    row,
    error: row === 4 ? "" : "fake scheduling failure"
  });
  sandbox.markSchedulingRowsStarted_ = () => {};
  const scheduled = sandbox.processOneQueuedJob_({
    action: "schedule", rows: [4, 5], attemptedRows: [4, 5],
    template: "Welcome", sender: "owner@akamai.com",
    scheduledForIso: "2099-01-01T14:00:00.000Z"
  });
  assert.deepEqual(cleared, [4, 5]);
  assert.equal(scheduled.successful, 1);
  assert.equal(scheduled.failed, 1);
});

test("sender authority is revalidated immediately before MailApp", () => {
  const sends = [];
  const sheet = { getRange() { return { setValue() {} }; } };
  const spreadsheet = { getSheetByName() { return {}; } };
  const sandbox = loadEmailsSandbox({ MailApp: { sendEmail: message => sends.push(message) } });
  sandbox.getTemplateStatusColumn_ = () => 2;
  sandbox.getHeaders_ = () => ({ status: 1 });
  sandbox.setStatus_ = () => {};
  sandbox.getRecipientNameForRow_ = () => "Recipient";
  sandbox.resolveRecipientsForRow_ = () => ({ to: "recipient@example.com", cc: "", bcc: "" });
  sandbox.getTemplateByKey_ = () => ({});
  sandbox.buildEmailPayload_ = () => ({
    subject: "Subject", plainBody: "Body", htmlBody: "<p>Body</p>",
    inlineImages: {}, attachments: [], senderName: "Owner", attachmentNames: ""
  });
  sandbox.stampTemplateFailure_ = () => {};
  sandbox.logSentEmail_ = () => {};
  sandbox.getAuthenticatedUserEmail_ = () => "other@akamai.com";
  const result = sandbox.sendOneRowNow_(
    spreadsheet, sheet, 2, { status: 1 }, "Welcome", "owner@akamai.com", {}
  );
  assert.equal(result.status, "failed");
  assert.equal(sends.length, 0);
});

test("dynamic action labels and duplicate-submit guard are present", () => {
  const script = formSource.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace("const context = <?!= contextJson ?>;", "const context = {};")
    .replace(/\s*initializeForm\(\);\s*$/, "");
  const sandbox = { Date, Intl, console, setTimeout() {}, google: {}, document: {} };
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox, { filename: "SendForm.html" });
  assert.equal(sandbox.getActionButtonText("send_now", 1, false), "Send 1 Email");
  assert.equal(sandbox.getActionButtonText("send_now", 4, false), "Send 4 Emails");
  assert.equal(sandbox.getActionButtonText("schedule", 1, false), "Schedule 1 Email");
  assert.equal(sandbox.getActionButtonText("schedule", 4, false), "Schedule 4 Emails");
  assert.equal(sandbox.getActionButtonText("schedule", 4, true), "Scheduling 4 Emails…");
  vm.runInContext(
    'context.actionMode="send_now"; context.senderState={status:"duplicate"}; selectedRowCount=1;',
    sandbox
  );
  assert.equal(sandbox.canSubmit(), false);
  vm.runInContext(
    'context.actionMode="schedule"; context.senderState={status:"ready"}; ' +
    'context.triggerState={status:"duplicate"}; selectedRowCount=1;',
    sandbox
  );
  assert.equal(sandbox.canSubmit(), false);
  assert.match(script, /if \(isSubmitting\) return;/);
  assert.match(script, /context\.senderState = result\.senderState;[\s\S]*renderSenderState\(\)/);
  assert.doesNotMatch(script, /saveInlineSenderProfile[\s\S]{0,1000}initializeForm\(\)/);
  assert.match(formSource, /id="senderProfileEmail" type="email" readonly/);
  assert.match(formSource, /id="senderOnboarding"/);
  assert.doesNotMatch(formSource, /<select id="sender">/);
  assert.match(formSource, /button-loading-sweep 1\.25s linear infinite/);
  assert.doesNotMatch(formSource, /id="senderDisplaySignature"/);
  assert.ok(
    script.indexOf('setSubmitButtonState(\n          "working"') <
    script.indexOf(".queueSendFormJob(formData)")
  );
});

test("repository contains no GmailApp and tests never use live triggers or real mail", () => {
  const projectSource = ["Code.js", "Data.js", "Emails.js", "Templates.js", "SendForm.html", "AddSender.html"]
    .map(file => fs.readFileSync(path.join(ROOT, file), "utf8"))
    .join("\n");
  assert.equal(projectSource.includes("GmailApp"), false);
  assert.equal(projectSource.includes("send as another user"), false);

  const addSenderScript = addSenderSource.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace("const context = <?!= contextJson ?>;", "const context = { imageAssets: [] };")
    .replace(/\s*initializeSenderForm\(\);\s*$/, "");
  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(addSenderScript, sandbox, { filename: "AddSender.html" });
});

let passed = 0;
tests.forEach(testCase => {
  try {
    testCase.fn();
    passed++;
    console.log("PASS " + testCase.name);
  } catch (err) {
    console.error("FAIL " + testCase.name);
    console.error(err.stack || err);
    process.exitCode = 1;
  }
});
console.log("\n" + passed + "/" + tests.length + " sender-onboarding tests passed.");
