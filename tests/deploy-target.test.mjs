import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import {
  assertProductionRelease,
  buildClaspInvocation,
  validateTarget
} from "../scripts/clasp-target.mjs";

const temp = fs.mkdtempSync(path.join(os.tmpdir(), "sendmebot-target-"));
const projectFile = path.join(temp, ".clasp.development.json");
fs.writeFileSync(projectFile, JSON.stringify({ scriptId: "development-id" }));

const targets = {
  "bound-dev": {
    environment: "development",
    projectFile: ".clasp.development.json",
    expectedScriptId: "development-id"
  }
};

const target = validateTarget(temp, "bound-dev", targets);
assert.equal(target.scriptId, "development-id");
assert.deepEqual(
  buildClaspInvocation("status", target),
  ["-P", projectFile, "status"]
);
assert.throws(() => validateTarget(temp, "missing", targets), /Unknown clasp target/);
assert.throws(() => validateTarget(temp, "bound-dev", {
  "bound-dev": Object.assign({}, targets["bound-dev"], { expectedScriptId: "wrong" })
}), /Script ID mismatch/);

const production = Object.assign({}, target, { alias: "bound-prod", environment: "production" });
assert.throws(
  () => assertProductionRelease(temp, production, ""),
  /confirm-production=bound-prod/
);

console.log("PASS named clasp targets reject ambiguity and mismatched IDs");
