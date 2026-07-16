#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

export const BOUND_PAYLOAD_FILES = [
  "Config.js",
  "Code.js",
  "Setup.js",
  "Data.js",
  "Emails.js",
  "Templates.js",
  "Modal.html",
  "SendForm.html",
  "SetupForm.html",
  "TemplateForm.html",
  "AddSender.html",
  "AddImage.html",
  "AssistantSidebar.html",
  "appsscript.json"
];

function apiFileFor(repoRoot, filename) {
  const extension = path.extname(filename);
  const name = extension === ".json" ? "appsscript" : path.basename(filename, extension);
  const type = extension === ".html" ? "HTML" : extension === ".json" ? "JSON" : "SERVER_JS";
  return {
    name,
    type,
    source: fs.readFileSync(path.join(repoRoot, filename), "utf8")
  };
}

export function canonicalizePayloadFiles(files) {
  return files.map(file => ({ name: file.name, type: file.type, source: file.source }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

export function payloadHash(files) {
  return crypto.createHash("sha256")
    .update(JSON.stringify(canonicalizePayloadFiles(files)))
    .digest("base64url");
}

export function buildPayload(options) {
  const repoRoot = options.repoRoot;
  const outputDir = options.outputDir;
  const configSource = fs.readFileSync(path.join(repoRoot, "Config.js"), "utf8");
  const versionMatch = configSource.match(/SENDMEBOT_APP_VERSION\s*=\s*["']([^"']+)["']/);
  if (!versionMatch) throw new Error("Could not determine SENDMEBOT_APP_VERSION.");
  const version = versionMatch[1];
  const files = canonicalizePayloadFiles(BOUND_PAYLOAD_FILES.map(file => apiFileFor(repoRoot, file)));
  const hash = payloadHash(files);
  const catalog = JSON.parse(fs.readFileSync(path.join(repoRoot, "release", "catalog.json"), "utf8"));
  const current = catalog.releases && catalog.releases[version];
  if (!current) throw new Error("Release catalog does not contain version " + version + ".");
  if (current.payloadSha256 !== "PENDING" && current.payloadSha256 !== hash) {
    throw new Error("Release catalog hash does not match canonical source for version " + version + ".");
  }

  fs.mkdirSync(outputDir, { recursive: true });
  const source = [
    "const SENDMEBOT_EMBEDDED_PAYLOAD = " + JSON.stringify({
      appId: "sendmebot",
      version,
      payloadSha256: hash,
      files
    }, null, 2) + ";",
    "const SENDMEBOT_RELEASE_CATALOG = " + JSON.stringify(catalog, null, 2) + ";",
    ""
  ].join("\n");
  fs.writeFileSync(path.join(outputDir, "Payload.js"), source);
  return { version, hash, outputFile: path.join(outputDir, "Payload.js") };
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  try {
    const repoRoot = path.resolve(import.meta.dirname, "..");
    const outputArg = process.argv.find(arg => arg.startsWith("--output="));
    const outputDir = outputArg
      ? path.resolve(repoRoot, outputArg.split("=")[1])
      : path.join(repoRoot, "installer", "dist", "development");
    const result = buildPayload({ repoRoot, outputDir });
    process.stdout.write(JSON.stringify(result) + "\n");
  } catch (err) {
    process.stderr.write((err.message || String(err)) + "\n");
    process.exitCode = 1;
  }
}
