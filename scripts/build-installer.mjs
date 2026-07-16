#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { buildPayload } from "./build-payload.mjs";

function parseArgs(argv) {
  const output = { environment: "development", config: "config/environments.local.json" };
  argv.forEach(arg => {
    if (arg.startsWith("--environment=")) output.environment = arg.split("=")[1];
    if (arg.startsWith("--config=")) output.config = arg.split("=")[1];
  });
  return output;
}

export function buildInstaller(options) {
  const repoRoot = options.repoRoot;
  const environmentName = options.environment;
  const configPath = path.resolve(repoRoot, options.config);
  if (!fs.existsSync(configPath)) throw new Error("Missing installer environment config: " + configPath);
  const environments = JSON.parse(fs.readFileSync(configPath, "utf8"));
  const environment = environments[environmentName];
  if (!environment) throw new Error("Unknown installer environment: " + environmentName);
  if (!environment.templateSpreadsheetId || !environment.releaseMetadataUrl) {
    throw new Error("Installer environment requires templateSpreadsheetId and releaseMetadataUrl.");
  }

  const sourceDir = path.join(repoRoot, "installer", "src");
  const outputDir = path.join(repoRoot, "installer", "dist", environmentName);
  fs.rmSync(outputDir, { recursive: true, force: true });
  fs.mkdirSync(outputDir, { recursive: true });

  fs.readdirSync(sourceDir)
    .filter(name => /\.(js|json)$/.test(name))
    .sort()
    .forEach(name => fs.copyFileSync(path.join(sourceDir, name), path.join(outputDir, name)));

  const generated = "const SENDMEBOT_INSTALLER_ENVIRONMENT = " +
    JSON.stringify(Object.assign({ name: environmentName }, environment), null, 2) + ";\n";
  fs.writeFileSync(path.join(outputDir, "Environment.js"), generated);
  buildPayload({ repoRoot, outputDir });
  return outputDir;
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  try {
    const args = parseArgs(process.argv.slice(2));
    const repoRoot = path.resolve(import.meta.dirname, "..");
    process.stdout.write(buildInstaller({ repoRoot, ...args }) + "\n");
  } catch (err) {
    process.stderr.write((err.message || String(err)) + "\n");
    process.exitCode = 1;
  }
}
