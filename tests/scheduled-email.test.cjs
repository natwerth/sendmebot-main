"use strict";

process.env.TZ = "America/Chicago";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const emailsSource = fs.readFileSync(
  path.join(__dirname, "..", "Emails.js"),
  "utf8"
);

const TODAY = new Date(2026, 6, 10, 12, 0, 0, 0);
let environment;

function formatMonthDay(value) {
  const date = new Date(value);
  return (date.getMonth() + 1) + "/" + date.getDate();
}

function createLock(canAcquire) {
  return {
    acquired: false,
    released: false,
    tryLockCalls: 0,
    tryLock: function(waitMilliseconds) {
      this.tryLockCalls++;
      this.waitMilliseconds = waitMilliseconds;
      this.acquired = canAcquire;
      return canAcquire;
    },
    releaseLock: function() {
      this.released = true;
    }
  };
}

function createScheduledRow(date, options) {
  const config = options || {};
  const sender = Object.prototype.hasOwnProperty.call(config, "sender")
    ? config.sender
    : "owner@example.com";
  const note = Object.prototype.hasOwnProperty.call(config, "note")
    ? config.note
    : JSON.stringify({
        template: "Welcome",
        sender: sender,
        scheduledDate: new Date(date).toISOString(),
        recipientConfig: {
          sender: sender,
          toField: { value: "Email", valueType: "field" },
          ccFields: [],
          bccFields: []
        }
      });

  return {
    statusValue: "Scheduled",
    templateValue: config.templateValue || "Scheduled for " + formatMonthDay(date),
    note: note,
    statusWrites: [],
    templateWrites: [],
    noteClears: 0
  };
}

function createSheet(rows) {
  return {
    rows: rows,
    dataReads: 0,
    getName: function() {
      return "Tracker";
    },
    getDataRange: function() {
      this.dataReads++;
      return {
        getValues: () => {
          return [["Status", "Template Email"]].concat(
            this.rows.map(row => [row.statusValue, row.templateValue])
          );
        }
      };
    },
    getRange: function(rowNumber, columnNumber) {
      const row = this.rows[rowNumber - 2];
      if (!row) throw new Error("Unknown fake row: " + rowNumber);

      if (columnNumber === 1) {
        return {
          getDisplayValue: function() {
            return row.statusValue;
          },
          getNote: function() {
            return "";
          },
          setValue: function(value) {
            row.statusValue = value;
            row.statusWrites.push(value);
            return this;
          },
          clearNote: function() {
            return this;
          }
        };
      }

      if (columnNumber === 2) {
        return {
          getDisplayValue: function() {
            return row.templateValue;
          },
          getNote: function() {
            return row.note;
          },
          setValue: function(value) {
            row.templateValue = value;
            row.templateWrites.push(value);
            return this;
          },
          clearNote: function() {
            row.note = "";
            row.noteClears++;
            return this;
          }
        };
      }

      throw new Error("Unknown fake column: " + columnNumber);
    }
  };
}

function resetEnvironment(rows) {
  environment = {
    sheet: createSheet(rows),
    templateSheet: {},
    lock: createLock(true),
    effectiveUser: "owner@example.com",
    sends: [],
    logs: [],
    sentLogRows: [],
    toasts: [],
    mailError: null
  };

  environment.spreadsheet = {
    getSheetByName: function(name) {
      return name === "Templates" ? environment.templateSheet : null;
    },
    toast: function(message, title, seconds) {
      environment.toasts.push({ message: message, title: title, seconds: seconds });
    }
  };

  return environment;
}

const sandbox = {
  console: console,
  Logger: {
    log: function(message) {
      environment.logs.push(String(message));
    }
  },
  LockService: {
    getScriptLock: function() {
      return environment.lock;
    },
    getDocumentLock: function() {
      return environment.lock;
    }
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: function() {
      return environment.spreadsheet;
    },
    flush: function() {}
  },
  Session: {
    getEffectiveUser: function() {
      return {
        getEmail: function() {
          return environment.effectiveUser;
        }
      };
    },
    getActiveUser: function() {
      return {
        getEmail: function() {
          return environment.effectiveUser;
        }
      };
    },
    getScriptTimeZone: function() {
      return "America/Chicago";
    }
  },
  Utilities: {
    formatDate: function(date, timeZone, pattern) {
      if (pattern === "M/d") return formatMonthDay(date);
      return formatMonthDay(date) + "/" + new Date(date).getFullYear();
    }
  },
  MailApp: {
    sendEmail: function(message) {
      if (environment.mailError) throw environment.mailError;
      environment.sends.push(message);
    }
  },
  getTrackerSheet_: function() {
    return environment.sheet;
  },
  getHeaders_: function() {
    return { status: 1, "template email": 2 };
  },
  getTemplateStatusColumns_: function() {
    return [{ headerName: "template email", col: 2 }];
  },
  normalize_: function(value) {
    return String(value || "").trim().toLowerCase();
  },
  getTemplateByKey_: function() {
    return { name: "Welcome" };
  },
  setStatus_: function(sheet, row, statusCol, value) {
    if (statusCol) sheet.getRange(row, statusCol).setValue(value);
  },
  stampTemplateColumn_: function(sheet, row) {
    sheet.getRange(row, 2).setValue("Sent on " + formatMonthDay(TODAY));
  },
  stampTemplateFailure_: function(sheet, row) {
    sheet.getRange(row, 2).setValue("Error: Not Sent");
  },
  logSentEmail_: function(ss, data) {
    environment.sentLogRows.push(data);
  }
};

vm.createContext(sandbox);
vm.runInContext(emailsSource, sandbox, { filename: "Emails.js" });

// Emails.js defines these helpers, so replace them after loading the file.
sandbox.getRecipientNameForRow_ = function() {
  return "Test Recipient";
};
sandbox.resolveRecipientsForRow_ = function() {
  return { to: "recipient@example.com", cc: "", bcc: "" };
};
sandbox.buildEmailPayload_ = function() {
  return {
    subject: "Subject",
    plainBody: "Plain body",
    htmlBody: "<p>HTML body</p>",
    inlineImages: {},
    attachments: [],
    attachmentNames: "",
    senderName: "Owner"
  };
};

function runWorker(overrides) {
  const runtime = Object.assign({
    lock: environment.lock,
    spreadsheet: environment.spreadsheet,
    sheet: environment.sheet,
    templateSheet: environment.templateSheet,
    effectiveUser: environment.effectiveUser,
    now: TODAY
  }, overrides || {});

  return sandbox.processScheduledEmails_(runtime);
}

const tests = [];

function test(name, fn) {
  tests.push({ name: name, fn: fn });
}

test("future scheduled dates are skipped", function() {
  const row = createScheduledRow(new Date(2026, 6, 11));
  resetEnvironment([row]);

  const summary = runWorker();

  assert.equal(summary.sent, 0);
  assert.equal(environment.sends.length, 0);
  assert.equal(row.templateValue, "Scheduled for 7/11");
  assert.notEqual(row.note, "");
});

test("a scheduled date today is processed", function() {
  const row = createScheduledRow(TODAY);
  resetEnvironment([row]);

  const summary = runWorker();

  assert.equal(summary.sent, 1);
  assert.equal(environment.sends.length, 1);
  assert.equal(row.statusValue, "Sent");
});

test("a past scheduled date is processed", function() {
  const row = createScheduledRow(new Date(2026, 6, 9));
  resetEnvironment([row]);

  const summary = runWorker();

  assert.equal(summary.sent, 1);
  assert.equal(environment.sends.length, 1);
});

test("trigger and non-boolean arguments never activate force mode", function() {
  const argumentsToTest = [
    { triggerUid: "123", hour: 8 },
    true,
    false,
    undefined,
    null,
    {},
    [],
    "true",
    "force",
    1,
    0,
    -1
  ];

  argumentsToTest.forEach(argument => {
    const row = createScheduledRow(new Date(2099, 0, 1));
    resetEnvironment([row]);

    const summary = sandbox.sendScheduledEmails(argument);

    assert.equal(summary.sent, 0);
    assert.equal(environment.sends.length, 0);
    assert.equal(row.templateValue, "Scheduled for 1/1");
  });
});

test("invalid JSON does not stop a later scheduled cell", function() {
  const brokenRow = createScheduledRow(TODAY, { note: "{not valid json" });
  const validRow = createScheduledRow(TODAY);
  resetEnvironment([brokenRow, validRow]);

  const summary = runWorker();

  assert.equal(brokenRow.templateValue, "Schedule Broken");
  assert.equal(brokenRow.note, "{not valid json");
  assert.equal(validRow.statusValue, "Sent");
  assert.equal(summary.sent, 1);
  assert.equal(summary.skipped, 1);
  assert.equal(environment.sends.length, 1);
});

test("a different trigger owner is skipped before visible status changes", function() {
  const row = createScheduledRow(TODAY, { sender: "other@example.com" });
  resetEnvironment([row]);

  const summary = runWorker();

  assert.equal(summary.sent, 0);
  assert.equal(summary.skipped, 1);
  assert.deepEqual(row.statusWrites, []);
  assert.deepEqual(row.templateWrites, []);
  assert.notEqual(row.note, "");
  assert.equal(environment.sends.length, 0);
});

test("an overlapping execution is rejected by the script lock", function() {
  const row = createScheduledRow(TODAY);
  resetEnvironment([row]);
  environment.lock = createLock(false);

  const summary = runWorker();

  assert.equal(summary.locked, true);
  assert.equal(environment.sheet.dataReads, 0);
  assert.equal(environment.sends.length, 0);
  assert.equal(environment.lock.waitMilliseconds, 30000);
  assert.equal(environment.lock.released, false);
});

test("a successful scheduled send cannot be sent again", function() {
  const row = createScheduledRow(TODAY);
  resetEnvironment([row]);

  const firstSummary = runWorker();
  environment.lock = createLock(true);
  const secondSummary = runWorker();

  assert.equal(firstSummary.sent, 1);
  assert.equal(secondSummary.sent, 0);
  assert.equal(environment.sends.length, 1);
  assert.equal(row.templateValue, "Sent on 7/10");
  assert.equal(row.note, "");
  assert.equal(row.noteClears, 1);
});

test("a failed send leaves an intelligible recoverable state", function() {
  const row = createScheduledRow(TODAY);
  const originalNote = row.note;
  resetEnvironment([row]);
  environment.mailError = new Error("Fake mail failure");

  const summary = runWorker();

  assert.equal(summary.failed, 1);
  assert.equal(row.templateValue, "Error: Not Sent");
  assert.equal(row.statusValue, "Failed");
  assert.equal(row.note, originalNote);
  assert.equal(environment.sentLogRows.length, 1);
  assert.equal(environment.sentLogRows[0].status, "Failed");
});

test("existing valid scheduled-note metadata remains readable", function() {
  const row = createScheduledRow(TODAY);
  resetEnvironment([row]);

  const data = sandbox.getScheduledEmailData_(environment.sheet, 2, 2);

  assert.equal(data.template, "Welcome");
  assert.equal(data.sender, "owner@example.com");
  assert.equal(data.recipientConfig.toField.value, "Email");
  assert.equal(data.scheduledDate, new Date(TODAY).toISOString());
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
