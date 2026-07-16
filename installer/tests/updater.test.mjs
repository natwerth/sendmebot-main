import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { BOUND_PAYLOAD_FILES, payloadHash } from "../../scripts/build-payload.mjs";

const ROOT = path.resolve(import.meta.dirname, "..");
const sources = ["Config.js", "Registry.js", "Updates.js"]
  .map(name => fs.readFileSync(path.join(ROOT, "src", name), "utf8"));

function digest(value) {
  return Array.from(crypto.createHash("sha256").update(String(value)).digest());
}

const sandbox = {
  console,
  Date,
  JSON,
  Utilities: {
    DigestAlgorithm: { SHA_256: "SHA_256" },
    computeDigest: (algorithm, value) => digest(value),
    base64EncodeWebSafe: value => Buffer.from(value).toString("base64url")
  }
};
vm.createContext(sandbox);
sources.forEach((source, index) => vm.runInContext(source, sandbox, { filename: String(index) }));

const manifest = scopes => ({
  name: "appsscript",
  type: "JSON",
  source: JSON.stringify({ oauthScopes: scopes })
});
const configMarker = source => ({ name: "Config", type: "SERVER_JS", source });

assert.equal(sandbox.hasNewManifestScopes_(
  [manifest(["scope-a"])],
  [manifest(["scope-a", "scope-b"])]
), true);
assert.equal(sandbox.hasNewManifestScopes_(
  [manifest(["scope-a"])],
  [manifest(["scope-a"])]
), false);

const payloadFiles = [
  configMarker('const SENDMEBOT_APP_ID = "sendmebot"; const version = "new";'),
  manifest(["scope-a"])
];
const nextHash = sandbox.hashAppsScriptFiles_(payloadFiles);
const oldFiles = [
  configMarker('const SENDMEBOT_APP_ID = "sendmebot"; const version = "old";'),
  manifest(["scope-a"])
];
const oldHash = sandbox.hashAppsScriptFiles_(oldFiles);
const payload = { version: "2.0.0", payloadSha256: nextHash, files: payloadFiles };
const catalog = { releases: { "1.0.0": { payloadSha256: oldHash } } };
const metadata = {
  channel: "stable",
  version: "2.0.0",
  payloadSha256: nextHash,
  updateMode: "automatic",
  requiresReauthorization: false,
  rolloutPercent: 100
};

sandbox.SpreadsheetApp = { openById: () => ({}) };
sandbox.readSendMeBotRegistry_ = () => ({
  appId: "sendmebot",
  spreadsheetId: "sheet-1",
  scriptId: "script-1"
});
sandbox.verifyUpdateTarget_ = () => ({ files: oldFiles });
sandbox.getEmbeddedPayload_ = () => payload;
sandbox.getEmbeddedReleaseCatalog_ = () => catalog;

const calls = [];
sandbox.requestAppsScriptApi_ = (method, requestPath, body) => {
  calls.push({ method, requestPath, body });
  if (method === "get") return { files: [configMarker("unexpected"), manifest(["scope-a"])] };
  return {};
};
assert.throws(
  () => sandbox.updateOneRegistration_({ spreadsheetId: "sheet-1", scriptId: "script-1" }, metadata),
  /previous source was restored/
);
assert.deepEqual(calls.map(call => call.method), ["put", "get", "put"]);
assert.deepEqual(calls[2].body.files, oldFiles);

sandbox.verifyUpdateTarget_ = () => ({ files: [configMarker("drift"), manifest(["scope-a"])] });
assert.throws(
  () => sandbox.updateOneRegistration_({ spreadsheetId: "sheet-1", scriptId: "script-1" }, metadata),
  /differs from every known SendMeBot release/
);

assert.equal(BOUND_PAYLOAD_FILES.some(file => /tests|\.clasp|installer/.test(file)), false);
const catalogFile = JSON.parse(fs.readFileSync(path.resolve(ROOT, "..", "release", "catalog.json"), "utf8"));
assert.notEqual(catalogFile.releases["1.0.0"].payloadSha256, "PENDING");
assert.equal(payloadHash(payloadFiles), crypto.createHash("sha256")
  .update(JSON.stringify(payloadFiles.slice().sort((a, b) => a.name.localeCompare(b.name))))
  .digest("base64url"));

const updateSource = sources.join("\n");
assert.doesNotMatch(updateSource, /SENDMEBOT_SOURCE_SCRIPT_ID|projects\/[^"']+\/content[^\n]+source/i);
assert.match(updateSource, /script\.google\.com\/home\/usersettings/);

console.log("PASS updater rejects drift and scope changes and restores source after verification failure");
