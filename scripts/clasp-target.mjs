#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";

const ALLOWED_COMMANDS = new Set(["status", "push", "pull", "versions", "deployments"]);

export function loadTargets(repoRoot, configPath = "config/targets.local.json") {
  const absolute = path.resolve(repoRoot, configPath);
  if (!fs.existsSync(absolute)) {
    throw new Error(
      "Missing local target configuration: " + absolute +
      ". Copy config/targets.example.json and fill in the IDs locally."
    );
  }
  return JSON.parse(fs.readFileSync(absolute, "utf8"));
}

export function validateTarget(repoRoot, alias, targets) {
  if (!alias) throw new Error("A named target is required.");
  const target = targets[alias];
  if (!target) throw new Error("Unknown clasp target: " + alias + ".");
  if (!target.projectFile || !target.expectedScriptId) {
    throw new Error("Target " + alias + " is missing projectFile or expectedScriptId.");
  }
  const projectFile = path.resolve(repoRoot, target.projectFile);
  if (!projectFile.startsWith(path.resolve(repoRoot) + path.sep)) {
    throw new Error("Target project files must remain inside the repository.");
  }
  if (!fs.existsSync(projectFile)) {
    throw new Error("Missing clasp project file for " + alias + ": " + projectFile);
  }
  const project = JSON.parse(fs.readFileSync(projectFile, "utf8"));
  if (project.scriptId !== target.expectedScriptId) {
    throw new Error("Script ID mismatch for " + alias + "; refusing to continue.");
  }
  return Object.assign({}, target, { alias, projectFile, scriptId: project.scriptId });
}

export function assertProductionRelease(repoRoot, target, confirmation, run = spawnSync) {
  if (target.environment !== "production") return;
  if (confirmation !== target.alias) {
    throw new Error(
      "Production target requires --confirm-production=" + target.alias + "."
    );
  }
  const status = run("git", ["status", "--porcelain"], {
    cwd: repoRoot,
    encoding: "utf8"
  });
  if (status.status !== 0 || String(status.stdout || "").trim()) {
    throw new Error("Production commands require a clean Git worktree.");
  }
  const tag = run("git", ["describe", "--tags", "--exact-match"], {
    cwd: repoRoot,
    encoding: "utf8"
  });
  if (tag.status !== 0 || !String(tag.stdout || "").trim()) {
    throw new Error("Production commands require the checked-out commit to have an exact tag.");
  }
}

export function buildClaspInvocation(command, target, extraArgs = []) {
  if (!ALLOWED_COMMANDS.has(command)) throw new Error("Unsupported clasp command: " + command + ".");
  return ["-P", target.projectFile, command].concat(extraArgs);
}

function parseArgs(argv) {
  const positional = [];
  const extra = [];
  let dryRun = false;
  let confirmation = "";
  let passthrough = false;
  argv.forEach(arg => {
    if (passthrough) {
      extra.push(arg);
    } else if (arg === "--") {
      passthrough = true;
    } else if (arg === "--dry-run") {
      dryRun = true;
    } else if (arg.startsWith("--confirm-production=")) {
      confirmation = arg.slice("--confirm-production=".length);
    } else {
      positional.push(arg);
    }
  });
  return { command: positional[0], alias: positional[1], dryRun, confirmation, extra };
}

export function main(argv = process.argv.slice(2), repoRoot = path.resolve(import.meta.dirname, "..")) {
  const parsed = parseArgs(argv);
  if (!parsed.command || !parsed.alias) {
    throw new Error(
      "Usage: node scripts/clasp-target.mjs <status|push|pull|versions|deployments> <target> [--dry-run]"
    );
  }
  const targets = loadTargets(repoRoot);
  const target = validateTarget(repoRoot, parsed.alias, targets);
  if (parsed.command === "push") {
    assertProductionRelease(repoRoot, target, parsed.confirmation);
  }
  const args = buildClaspInvocation(parsed.command, target, parsed.extra);
  if (parsed.dryRun) {
    if (parsed.command === "push" && target.kind === "addon") {
      process.stdout.write(
        "node scripts/build-installer.mjs --environment=" + target.environment + "\n"
      );
    }
    process.stdout.write("clasp " + args.join(" ") + "\n");
    return 0;
  }
  if (parsed.command === "push" && target.kind === "addon") {
    const build = spawnSync(
      process.execPath,
      ["scripts/build-installer.mjs", "--environment=" + target.environment],
      { cwd: repoRoot, stdio: "inherit" }
    );
    if (build.error) throw build.error;
    if (build.status) return build.status;
  }
  const result = spawnSync("clasp", args, { cwd: repoRoot, stdio: "inherit" });
  if (result.error) throw result.error;
  return result.status || 0;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    process.exitCode = main();
  } catch (err) {
    process.stderr.write((err && err.message ? err.message : String(err)) + "\n");
    process.exitCode = 1;
  }
}
