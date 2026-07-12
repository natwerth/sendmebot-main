"use strict";

process.env.TZ = "America/Chicago";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const emailsSource = fs.readFileSync(path.join(ROOT, "Emails.js"), "utf8");
const dataSource = fs.readFileSync(path.join(ROOT, "Data.js"), "utf8");
const formSource = fs.readFileSync(path.join(ROOT, "SendForm.html"), "utf8");
const NOW = new Date("2026-07-10T17:00:00.000Z");

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

function createLock(canAcquire = true) {
  return {
    released: false,
    tryLockCalls: 0,
    tryLock(milliseconds) {
      this.tryLockCalls++;
      this.waitMilliseconds = milliseconds;
      return canAcquire;
    },
    releaseLock() { this.released = true; }
  };
}

function cloneCell(value) {
  return value instanceof Date ? new Date(value.getTime()) : value;
}

function createGridSheet(name, headers, rows) {
  const sheet = {
    name,
    headers: headers.slice(),
    rows: rows.map(row => row.map(cloneCell)),
    notes: rows.map(row => row.map(() => "")),
    writes: [],
    notesWritten: 0,
    getName() { return this.name; },
    getLastRow() { return this.rows.length + 1; },
    getLastColumn() { return this.headers.length; },
    getDataRange() {
      return {
        getValues: () => [this.headers.slice()].concat(this.rows.map(row => row.map(cloneCell))),
        getDisplayValues: () => [this.headers.slice()].concat(this.rows.map(row => row.map(value => String(value || "")))),
        getNotes: () => [this.headers.map(() => "")].concat(this.notes.map(row => row.slice()))
      };
    },
    getRange(row, col, numRows, numCols) {
      const target = this;
      return {
        setValue(value) {
          if (row === 1) target.headers[col - 1] = value;
          else target.rows[row - 2][col - 1] = value;
          target.writes.push({ row, col, value });
          return this;
        },
        setValues(values) {
          if (row === 1) target.headers = values[0].slice();
          else target.rows[row - 2] = values[0].slice();
          target.writes.push({ row, col, values: values[0].slice() });
          return this;
        },
        getDisplayValue() {
          const value = row === 1 ? target.headers[col - 1] : target.rows[row - 2][col - 1];
          return String(value || "");
        },
        getValue() {
          return row === 1 ? target.headers[col - 1] : target.rows[row - 2][col - 1];
        },
        getValues() {
          if (row === 1) return [target.headers.slice(col - 1, col - 1 + (numCols || 1))];
          return [target.rows[row - 2].slice(col - 1, col - 1 + (numCols || 1))];
        },
        getDisplayValues() {
          const values = [];
          for (let r = 0; r < (numRows || 1); r++) {
            const source = row + r === 1 ? target.headers : target.rows[row + r - 2];
            values.push(source.slice(col - 1, col - 1 + (numCols || 1)).map(value => String(value || "")));
          }
          return values;
        },
        getA1Notation() { return "R" + row + "C" + col; },
        setNote() { target.notesWritten++; return this; },
        clearNote() {
          if (row > 1) target.notes[row - 2][col - 1] = "";
          return this;
        }
      };
    }
  };
  return sheet;
}

function headersFor(sheet) {
  const output = {};
  sheet.headers.forEach((header, index) => {
    const key = String(header || "").trim().toLowerCase();
    if (key) output[key] = index + 1;
  });
  return output;
}

const SENT_HEADERS = [
  "Timestamp", "Status", "Scheduled For", "Processed At", "Message", "Name",
  "Recipient", "Sender", "CC", "BCC", "Template", "Subject", "Email Body",
  "Attachments", "Log Note"
];

function metadata(overrides) {
  return JSON.stringify(Object.assign({
    version: 3,
    sourceSheet: "Tracker",
    sourceRow: 2,
    sourceStatusColumn: 2,
    recordIdHeader: "Student Name",
    recordIdValue: "Alice Student",
    scheduledTimeZone: "America/Chicago",
    scheduledDisplayText: "Jul 10, 2026, 11:00 AM CDT"
  }, overrides || {}));
}

function makeSentRow(scheduledFor, overrides, headers = SENT_HEADERS) {
  const values = {
    "Timestamp": new Date("2026-07-09T15:00:00.000Z"),
    "Status": "Scheduled",
    "Scheduled For": scheduledFor,
    "Processed At": "",
    "Message": "Email scheduled.",
    "Name": "Recipient Name",
    "Recipient": "recipient@example.com",
    "Sender": "owner@example.com",
    "CC": "",
    "BCC": "",
    "Template": "Welcome",
    "Subject": "Frozen subject",
    "Email Body": "Frozen plain body",
    "Attachments": "",
    "Log Note": metadata()
  };
  Object.assign(values, overrides || {});
  return headers.map(header => values[header]);
}

function createServerEnvironment(sentRows, sentHeaders = SENT_HEADERS) {
  const sentSheet = createGridSheet("Sent", sentHeaders, sentRows);
  const tracker = createGridSheet(
    "Tracker",
    ["Status", "Welcome", "Student Name"],
    [["Scheduled", "Scheduled for 7/10", "Alice Student"]]
  );
  const templateSheet = {};
  const environment = {
    sentSheet,
    tracker,
    templateSheet,
    lock: createLock(),
    sends: [],
    logs: [],
    queuedLogs: [],
    logError: null,
    effectiveUser: "owner@example.com"
  };

  environment.spreadsheet = {
    getSpreadsheetTimeZone() { return "America/Chicago"; },
    getSheetByName(name) {
      if (name === "Sent") return sentSheet;
      if (name === "Tracker") return tracker;
      if (name === "Templates") return templateSheet;
      return null;
    },
    toast() {}
  };

  const sandbox = {
    Date,
    JSON,
    console,
    Logger: { log(message) { environment.logs.push(String(message)); } },
    LockService: {
      getScriptLock() { return environment.lock; },
      getDocumentLock() { return environment.lock; }
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() { return environment.spreadsheet; },
      flush() {}
    },
    Session: {
      getActiveUser() { return { getEmail: () => environment.effectiveUser }; },
      getEffectiveUser() { return { getEmail: () => environment.effectiveUser }; },
      getScriptTimeZone() { return "America/Chicago"; }
    },
    Utilities: {
      formatDate(date, zone, pattern) {
        if (pattern === "M/d") return (date.getMonth() + 1) + "/" + date.getDate();
        return "Jul 10, 2026 11:00 AM CDT";
      }
    },
    MailApp: { sendEmail(message) { environment.sends.push(message); } },
    getHeaders_: headersFor,
    normalize_(value) { return String(value || "").trim().toLowerCase(); },
    normalizeTemplateKey_(value) { return String(value || "").trim().toLowerCase(); },
    getAuthenticatedUserEmail_() { return environment.effectiveUser; },
    getTrackerSheet_() { return tracker; },
    getTemplateStatusColumn_() { return 2; },
    getTemplateByKey_() { return { attachmentLink: "" }; },
    getSenderProfile_() { return { name: "Owner", email: "owner@example.com", signatureText: "" }; },
    getRowData_() { return { "Student Name": "Alice Student" }; },
    resolveRecipientsForRow_() { return { to: "recipient@example.com", cc: "", bcc: "" }; },
    getRecipientNameForRow_() { return "Recipient Name"; },
    buildEmailPayload_() {
      return {
        subject: "Frozen subject",
        plainBody: "Frozen plain body",
        htmlBody: "<p>Frozen HTML</p>",
        inlineImages: {},
        attachments: [],
        attachmentNames: "",
        senderName: "Owner"
      };
    },
    getImageAssets_() { return {}; },
    stampTemplateColumn_(sheet, row) { sheet.getRange(row, 2).setValue("Sent on 7/10"); },
    stampTemplateFailure_(sheet, row) { sheet.getRange(row, 2).setValue("Error: Not Sent"); },
    logSentEmail_(ss, logData) {
      if (environment.logError) throw environment.logError;
      environment.queuedLogs.push(logData);
      return 2;
    }
  };

  vm.createContext(sandbox);
  vm.runInContext(emailsSource, sandbox, { filename: "Emails.js" });
  // Replace helpers that Emails.js defines itself and which are external in Apps Script.
  sandbox.getHeaders_ = headersFor;
  sandbox.normalize_ = value => String(value || "").trim().toLowerCase();
  sandbox.normalizeTemplateKey_ = value => String(value || "").trim().toLowerCase();
  sandbox.getAuthenticatedUserEmail_ = () => environment.effectiveUser;
  sandbox.getTrackerSheet_ = () => tracker;
  sandbox.getTemplateStatusColumn_ = () => 2;
  sandbox.getTemplateByKey_ = () => ({ attachmentLink: "" });
  sandbox.getSenderProfile_ = () => ({ name: "Owner", email: "owner@example.com", signatureText: "" });
  sandbox.getRowData_ = () => ({ "Student Name": "Alice Student" });
  sandbox.resolveRecipientsForRow_ = () => ({ to: "recipient@example.com", cc: "", bcc: "" });
  sandbox.getRecipientNameForRow_ = () => "Recipient Name";
  sandbox.buildEmailPayload_ = () => ({
    subject: "Frozen subject", plainBody: "Frozen plain body", htmlBody: "<p>Frozen HTML</p>",
    inlineImages: {}, attachments: [], attachmentNames: "", senderName: "Owner"
  });
  sandbox.getImageAssets_ = () => ({});
  sandbox.stampTemplateColumn_ = (sheet, row) => sheet.getRange(row, 2).setValue("Sent on 7/10");
  sandbox.stampTemplateFailure_ = (sheet, row) => sheet.getRange(row, 2).setValue("Error: Not Sent");
  sandbox.logSentEmail_ = (ss, logData) => {
    if (environment.logError) throw environment.logError;
    environment.queuedLogs.push(logData);
    return 2;
  };

  return { environment, sandbox };
}

function loadClientFunctions() {
  const script = formSource.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace("const context = <?!= contextJson ?>;", "const context = {};")
    .replace(/\s*initializeForm\(\);\s*$/, "");
  const elements = {};
  const sandbox = {
    Date, Intl, console, setTimeout() {}, google: {},
    document: {
      getElementById(id) {
        if (!elements[id]) elements[id] = { value: "", textContent: "", classList: { add() {}, remove() {} } };
        return elements[id];
      }
    },
    __elements: elements
  };
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox, { filename: "SendForm.html" });
  return sandbox;
}

test("tomorrow uses local date components and time defaults to 08:03", () => {
  const client = loadClientFunctions();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  client.initializeScheduleFields();
  assert.equal(client.__elements.scheduleDate.value, client.formatLocalDateInputValue(tomorrow));
  assert.equal(client.__elements.scheduleTime.value, "08:03");
  assert.match(formSource, /getFullYear\(\).*getMonth\(\).*getDate\(\)/s);
  assert.doesNotMatch(formSource, /toISOString\(\)\.slice\(0,\s*10\)/);
  assert.match(formSource, /scheduleTime"\)\.value = "08:03"/);
});

test("browser timezone is detected and included in a valid ISO payload", () => {
  const client = loadClientFunctions();
  const result = client.createSchedulePayload("2026-07-11", "08:03", NOW);
  assert.equal(client.detectBrowserTimeZone(), "America/Chicago");
  assert.equal(result.scheduledTimeZone, "America/Chicago");
  assert.equal(result.scheduledForIso, "2026-07-11T13:03:00.000Z");
  assert.ok(result.scheduledDisplayText);
});

test("past browser-local datetime is rejected", () => {
  const client = loadClientFunctions();
  assert.throws(
    () => client.createSchedulePayload("2026-07-10", "11:59", new Date("2026-07-10T17:00:00.000Z")),
    /future/
  );
});

test("lean scheduling freezes recipients and record identity without rendering", () => {
  const sheet = createGridSheet(
    "Tracker",
    ["Select", "Student Name", "Email", "Status", "Welcome"],
    [[true, "Alice Student", "alice@example.com", "", "Scheduling..."]]
  );
  const sandbox = createServerEnvironment([]).sandbox;
  sandbox.resolveRecipientsForRow_ = () => ({
    to: "alice@example.com", cc: "manager@example.com", bcc: "audit@example.com"
  });
  sandbox.buildEmailPayload_ = () => { throw new Error("rendering must not run while scheduling"); };
  const result = sandbox.prepareScheduledRows_(
    {}, sheet, [2], headersFor(sheet), {}, "Student Name"
  );
  assert.deepEqual(Array.from(result.validRows), [2]);
  assert.equal(result.failures.length, 0);
  assert.equal(result.preparedRows[2].recordIdValue, "Alice Student");
  assert.equal(result.preparedRows[2].recipients.cc, "manager@example.com");
  assert.equal(result.preparedRows[2].recipients.bcc, "audit@example.com");
  const source = emailsSource.slice(
    emailsSource.indexOf("function prepareScheduledRows_"),
    emailsSource.indexOf("function queueSendFormJob")
  );
  assert.doesNotMatch(source, /buildEmailPayload_|DriveApp|getTemplateByKey_/);
});

test("blank and case-insensitive duplicate record IDs fail row validation", () => {
  const sheet = createGridSheet(
    "Tracker",
    ["Select", "Student Name", "Email"],
    [
      [true, "Alice Student", "a@example.com"],
      [true, " alice student ", "b@example.com"],
      [true, "", "c@example.com"]
    ]
  );
  const sandbox = createServerEnvironment([]).sandbox;
  sandbox.resolveRecipientsForRow_ = () => ({ to: "recipient@example.com", cc: "", bcc: "" });
  const result = sandbox.prepareScheduledRows_(
    {}, sheet, [2, 3, 4], headersFor(sheet), {}, "Student Name"
  );
  assert.equal(result.validRows.length, 0);
  assert.equal(result.failures.length, 3);
  assert.match(result.failures[0].error, /not unique/);
  assert.match(result.failures[2].error, /blank/);
});

test("delivery token validation names missing columns but permits blank existing values", () => {
  const sandbox = createServerEnvironment([]).sandbox;
  const sender = { name: "Owner", email: "owner@example.com", signatureText: "" };
  assert.doesNotThrow(() => sandbox.validateLiveRenderingTokens_(
    { subject: "Hello {{Start Date}}", body: "{{Sender Signature}}" },
    sender,
    { "Start Date": "" },
    {}
  ));
  assert.throws(() => sandbox.validateLiveRenderingTokens_(
    { subject: "Hello {{Renamed Column}}", body: "" }, sender, {}, {}
  ), /Renamed Column.*no current Tracker column/i);
});

test("scheduling writes one complete queue record before tracker scheduled status", () => {
  const { environment, sandbox } = createServerEnvironment([]);
  environment.tracker.rows[0] = ["", "Scheduling...", "Alice Student"];
  const result = sandbox.scheduleOneRow_(
    environment.spreadsheet, environment.tracker, 2, headersFor(environment.tracker),
    "Welcome", "owner@example.com", new Date("2099-07-11T13:03:00.000Z"),
    "America/Chicago", "Jul 11, 2099, 8:03 AM CDT", {}, {
      recordIdHeader: "Student Name", recordIdValue: "Alice Student",
      name: "Alice Student", recipients: { to: "recipient@example.com", cc: "", bcc: "" }
    }
  );

  assert.equal(result.status, "scheduled");
  assert.equal(environment.queuedLogs.length, 1);
  const row = environment.queuedLogs[0];
  assert.ok(row.timestamp instanceof Date);
  assert.equal(row.status, "Scheduled");
  assert.equal(row.scheduledFor.toISOString(), "2099-07-11T13:03:00.000Z");
  assert.equal(row.processedAt, "");
  ["message", "name", "email", "sender", "template", "subject", "body", "attachments", "logNote"]
    .forEach(key => assert.ok(Object.prototype.hasOwnProperty.call(row, key), key));
  assert.equal(row.subject, "");
  assert.equal(row.body, "");
  assert.equal(JSON.parse(row.logNote).version, 3);
  assert.equal(environment.tracker.rows[0][1], "Scheduled for 7/11");
  assert.equal(environment.tracker.rows[0][0], "");
  assert.equal(environment.tracker.notesWritten, 0);
});

test("tracker never shows scheduled when the Sent write fails", () => {
  const { environment, sandbox } = createServerEnvironment([]);
  environment.tracker.rows[0] = ["", "Scheduling...", "Alice Student"];
  environment.logError = new Error("fake Sent write failure");
  const result = sandbox.scheduleOneRow_(
    environment.spreadsheet, environment.tracker, 2, headersFor(environment.tracker),
    "Welcome", "owner@example.com", new Date("2099-07-11T13:03:00.000Z"),
    "America/Chicago", "Jul 11, 2099, 8:03 AM CDT", {}, {
      recordIdHeader: "Student Name", recordIdValue: "Alice Student",
      name: "Alice Student", recipients: { to: "recipient@example.com", cc: "", bcc: "" }
    }
  );
  assert.equal(result.status, "failed");
  assert.match(environment.tracker.rows[0][1], /^Error: fake Sent write failure/);
  assert.equal(environment.tracker.writes.some(write => String(write.value).startsWith("Scheduled for ")), false);
  assert.equal(environment.tracker.notesWritten, 0);
});

test("future Scheduled row is skipped", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T17:01:00.000Z"))
  ]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  assert.equal(summary.sent, 0);
  assert.equal(summary.skipped, 1);
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.sentSheet.rows[0][1], "Scheduled");
});

test("due row renders current Tracker data and finalizes as Sent", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  environment.tracker.rows[0][1] = "Scheduled for 1/1";
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message), getFileBlob: id => ({ id })
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.sent, 1);
  assert.equal(environment.sends.length, 1);
  assert.equal(environment.sends[0].subject, "Frozen subject");
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Sent");
  assert.ok(environment.sentSheet.rows[0][headers["processed at"] - 1] instanceof Date);
  assert.equal(environment.sentSheet.rows[0][headers.subject - 1], "Frozen subject");
  assert.equal(environment.sentSheet.rows[0][headers["email body"] - 1], "Frozen plain body");
  assert.match(environment.sentSheet.rows[0][headers["log note"] - 1], /renderedHtmlBody/);
  assert.equal(environment.sentSheet.writes.some(write => write.value === "Processing"), true);
  assert.equal(environment.tracker.rows[0][1], "Sent on 7/10");
  assert.equal(environment.tracker.rows[0][0], "Scheduled");
});

test("blank tracker status cancels the due queue row without sending", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  environment.tracker.rows[0][1] = "";
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.cancelled, 1);
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Cancelled");
  assert.ok(environment.sentSheet.rows[0][headers["processed at"] - 1] instanceof Date);
  assert.match(environment.sentSheet.rows[0][headers.message - 1], /cancelled/i);
  assert.equal(environment.sends.length, 0);
});

test("changed tracker status cancels that row and processing continues", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:58:00.000Z")),
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {
      "Log Note": metadata({ sourceRow: 3, recordIdValue: "Bob Student" })
    })
  ]);
  environment.tracker.rows[0][1] = "Sent on 7/9";
  environment.tracker.rows.push(["Scheduled", "Scheduled for 7/10", "Bob Student"]);
  environment.tracker.notes.push(["", "", ""]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.cancelled, 1);
  assert.equal(summary.sent, 1);
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Cancelled");
  assert.ok(environment.sentSheet.rows[0][headers["processed at"] - 1] instanceof Date);
  assert.equal(environment.sentSheet.rows[1][headers.status - 1], "Sent");
  assert.equal(environment.sends.length, 1);
});

test("missing source sheet, record, or ID column orphans the due queue row without sending", () => {
  [
    { metadata: { sourceSheet: "Deleted Tracker" } },
    { metadata: { recordIdValue: "Deleted Student" } },
    { metadata: { recordIdHeader: "Deleted ID" } }
  ].forEach(testCase => {
    const { environment, sandbox } = createServerEnvironment([
      makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {
        "Log Note": metadata(testCase.metadata)
      })
    ]);
    const summary = sandbox.processScheduledEmails_({
      spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
      lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
      sendEmail: message => environment.sends.push(message)
    });
    const headers = headersFor(environment.sentSheet);
    assert.equal(summary.orphaned, 1);
    assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Orphaned");
    assert.ok(environment.sentSheet.rows[0][headers["processed at"] - 1] instanceof Date);
    assert.match(environment.sentSheet.rows[0][headers.message - 1], /no longer exists|could not be found/i);
    assert.equal(environment.sends.length, 0);
  });
});

test("failed due row is finalized and does not stop a later row", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:58:00.000Z"), { Recipient: "first@example.com" }),
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {
      Recipient: "second@example.com",
      "Log Note": metadata({ sourceRow: 3, recordIdValue: "Bob Student" })
    })
  ]);
  environment.tracker.rows.push(["Scheduled", "Scheduled for 7/10", "Bob Student"]);
  environment.tracker.notes.push(["", "", ""]);
  let calls = 0;
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail(message) {
      calls++;
      if (calls === 1) throw new Error("fake delivery failure");
      environment.sends.push(message);
    }
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.failed, 1);
  assert.equal(summary.sent, 1);
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Failed");
  assert.ok(environment.sentSheet.rows[0][headers["processed at"] - 1] instanceof Date);
  assert.equal(environment.sentSheet.rows[1][headers.status - 1], "Sent");
});

test("Sent and Processing rows cannot send twice", () => {
  const sent = makeSentRow(new Date("2026-07-10T16:00:00.000Z"), { Status: "Sent" });
  const processing = makeSentRow(new Date("2026-07-10T16:00:00.000Z"), { Status: "Processing" });
  const { environment, sandbox } = createServerEnvironment([sent, processing]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  assert.equal(summary.sent, 0);
  assert.equal(environment.sends.length, 0);
});

test("trigger owner skips another user's queue row without modifying it", () => {
  const otherRow = makeSentRow(new Date("2026-07-10T16:58:00.000Z"), {
    Sender: "other@example.com"
  });
  const ownRow = makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {
    "Log Note": metadata({ sourceRow: 3, recordIdValue: "Bob Student" })
  });
  const { environment, sandbox } = createServerEnvironment([otherRow, ownRow]);
  environment.tracker.rows.push(["Scheduled", "Scheduled for 7/10", "Bob Student"]);
  environment.tracker.notes.push(["", "", ""]);

  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: "owner@example.com",
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Scheduled");
  assert.equal(environment.sentSheet.rows[0][headers["processed at"] - 1], "");
  assert.equal(environment.sentSheet.rows[1][headers.status - 1], "Sent");
  assert.equal(summary.sent, 1);
  assert.equal(environment.sends.length, 1);
});

test("missing effective trigger identity processes nothing", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: "",
    sendEmail: message => environment.sends.push(message)
  });
  assert.equal(summary.sent, 0);
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.sentSheet.rows[0][1], "Scheduled");
});

test("scheduled sender ownership is revalidated immediately before delivery", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  let identityReads = 0;
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW,
    getEffectiveUserEmail() {
      identityReads++;
      return identityReads === 1 ? "owner@example.com" : "other@example.com";
    },
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.failed, 1);
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.sentSheet.rows[0][headers.status - 1], "Failed");
});

test("trigger event objects cannot bypass future-time validation", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2099-01-01T00:00:00.000Z"))
  ]);
  sandbox.sendScheduledEmails({ triggerUid: "123", force: true });
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.sentSheet.rows[0][1], "Scheduled");
});

test("script lock blocks overlapping processing", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  environment.lock = createLock(false);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW
  });
  assert.equal(summary.locked, true);
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.lock.released, false);
});

test("processor resolves shuffled Sent columns by header name", () => {
  const shuffled = [
    "Recipient", "Log Note", "Status", "Email Body", "Scheduled For", "Sender",
    "Processed At", "Subject", "Message", "Template", "CC", "BCC", "Timestamp",
    "Name", "Attachments"
  ];
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {}, shuffled)
  ], shuffled);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  assert.equal(summary.sent, 1);
  assert.equal(environment.sends[0].to, "recipient@example.com");
  assert.equal(environment.sentSheet.rows[0][shuffled.indexOf("Status")], "Sent");
});

test("version-2 scheduled rows fail clearly without sending", () => {
  const legacyMetadata = JSON.stringify({ version: 2, sourceSheet: "Tracker", sourceRow: 2 });
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"), { "Log Note": legacyMetadata })
  ]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.failed, 1);
  assert.equal(environment.sends.length, 0);
  assert.match(environment.sentSheet.rows[0][headers.message - 1], /version 2.*Reschedule/i);
});

test("record lookup follows a sorted row using case-insensitive ID matching", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"), {
      "Log Note": metadata({ sourceRow: 99, recordIdValue: "alice student" })
    })
  ]);
  environment.tracker.rows.unshift(["", "", "Someone Else"]);
  environment.tracker.notes.unshift(["", "", ""]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  assert.equal(summary.sent, 1);
  assert.equal(environment.tracker.rows[1][1], "Sent on 7/10");
});

test("duplicate record IDs fail closed", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  environment.tracker.rows.push(["Scheduled", "Scheduled for 7/10", "alice student"]);
  environment.tracker.notes.push(["", "", ""]);
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.failed, 1);
  assert.equal(environment.sends.length, 0);
  assert.match(environment.sentSheet.rows[0][headers.message - 1], /multiple Tracker rows/i);
});

test("missing delivery-time template variable writes detailed Sent and Tracker errors", () => {
  const { environment, sandbox } = createServerEnvironment([
    makeSentRow(new Date("2026-07-10T16:59:00.000Z"))
  ]);
  sandbox.getTemplateByKey_ = () => ({ subject: "Hello {{Deleted Column}}", body: "Body", attachmentLink: "" });
  sandbox.getRowData_ = () => ({ "Student Name": "Alice Student" });
  const summary = sandbox.processScheduledEmails_({
    spreadsheet: environment.spreadsheet, sentSheet: environment.sentSheet,
    lock: environment.lock, now: NOW, effectiveUser: environment.effectiveUser,
    sendEmail: message => environment.sends.push(message)
  });
  const headers = headersFor(environment.sentSheet);
  assert.equal(summary.failed, 1);
  assert.equal(environment.sends.length, 0);
  assert.match(environment.sentSheet.rows[0][headers.message - 1], /Deleted Column/);
  assert.match(environment.tracker.rows[0][1], /^Error: .*Deleted Column/);
});

test("manual scheduled-trigger setup uses one five-minute handler", () => {
  const setupSource = emailsSource.slice(
    emailsSource.indexOf("function getCurrentUserScheduledTriggerState_"),
    emailsSource.indexOf("function setupJobProcessorTrigger")
  );
  assert.match(setupSource, /getHandlerFunction\(\) === "sendScheduledEmails"/);
  assert.match(setupSource, /\.everyMinutes\(5\)/);
  assert.doesNotMatch(setupSource, /everyDays|atHour/);
});

test("immediate-send logging retains existing values with new queue columns blank", () => {
  const logSheet = createGridSheet("Sent", SENT_HEADERS, []);
  logSheet.insertRowBefore = function(row) { this.rows.splice(row - 2, 0, this.headers.map(() => "")); };
  const sandbox = {
    Date, JSON, console,
    Utilities: { formatDate() { return "7/10/2026 12:00:00 PM"; } },
    Session: { getScriptTimeZone() { return "America/Chicago"; } },
    normalize_: value => String(value || "").trim().toLowerCase(),
    getHeaders_: headersFor,
    SpreadsheetApp: {}, DriveApp: {}, Logger: { log() {} }
  };
  vm.createContext(sandbox);
  vm.runInContext(dataSource, sandbox, { filename: "Data.js" });
  sandbox.normalize_ = value => String(value || "").trim().toLowerCase();
  sandbox.getHeaders_ = headersFor;
  sandbox.logSentEmail_({ getSheetByName: () => logSheet }, {
    status: "Sent", message: "Email sent successfully.", name: "A", email: "a@example.com",
    sender: "owner@example.com", template: "Welcome", subject: "Hi", body: "Body"
  });
  const headers = headersFor(logSheet);
  assert.equal(logSheet.rows[0][headers.status - 1], "Sent");
  assert.equal(logSheet.rows[0][headers["scheduled for"] - 1], "");
  assert.equal(logSheet.rows[0][headers["processed at"] - 1], "");
  assert.equal(logSheet.rows[0][headers.message - 1], "Email sent successfully.");
  assert.equal(logSheet.rows[0][headers.recipient - 1], "a@example.com");
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
console.log("\n" + passed + "/" + tests.length + " scheduled-email tests passed.");
