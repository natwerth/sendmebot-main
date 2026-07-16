function canonicalizeAppsScriptFiles_(files) {
  return (files || []).map(file => ({
    name: String(file.name || ""),
    type: String(file.type || ""),
    source: String(file.source || "")
  })).sort((a, b) => a.name.localeCompare(b.name));
}


function sha256WebSafe_(value) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ""))
  ).replace(/=+$/, "");
}


function hashAppsScriptFiles_(files) {
  return sha256WebSafe_(JSON.stringify(canonicalizeAppsScriptFiles_(files)));
}


function getEmbeddedPayload_() {
  if (typeof SENDMEBOT_EMBEDDED_PAYLOAD === "undefined" || !SENDMEBOT_EMBEDDED_PAYLOAD) {
    throw new Error("The installer update payload was not generated.");
  }
  return SENDMEBOT_EMBEDDED_PAYLOAD;
}


function getEmbeddedReleaseCatalog_() {
  if (typeof SENDMEBOT_RELEASE_CATALOG === "undefined" || !SENDMEBOT_RELEASE_CATALOG) {
    throw new Error("The installer release catalog was not generated.");
  }
  return SENDMEBOT_RELEASE_CATALOG;
}


function fetchReleaseMetadata_() {
  const environment = getInstallerEnvironment_();
  const response = UrlFetchApp.fetch(environment.releaseMetadataUrl, { muteHttpExceptions: true });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("Release metadata request failed with HTTP " + status + ".");
  }
  return JSON.parse(response.getContentText());
}


function getManifestScopes_(files) {
  const manifest = (files || []).find(file => file.name === "appsscript");
  if (!manifest) return [];
  try {
    const parsed = JSON.parse(manifest.source || "{}");
    return Array.isArray(parsed.oauthScopes) ? parsed.oauthScopes.map(String).sort() : [];
  } catch (err) {
    throw new Error("The target Apps Script manifest is invalid JSON.");
  }
}


function hasNewManifestScopes_(currentFiles, nextFiles) {
  const current = getManifestScopes_(currentFiles);
  return getManifestScopes_(nextFiles).some(scope => current.indexOf(scope) === -1);
}


function isKnownReleaseHash_(hash, catalog) {
  const releases = catalog && catalog.releases ? catalog.releases : {};
  return Object.keys(releases).some(version => releases[version].payloadSha256 === hash);
}


function isTargetInRollout_(spreadsheetId, percentage) {
  const limit = Math.max(0, Math.min(100, Number(percentage === undefined ? 100 : percentage)));
  let hash = 0;
  String(spreadsheetId || "").split("").forEach(character => {
    hash = ((hash << 5) - hash + character.charCodeAt(0)) | 0;
  });
  return Math.abs(hash) % 100 < limit;
}


function compareSemanticVersions_(left, right) {
  const parse = value => String(value || "0").split(".").map(part => parseInt(part, 10) || 0);
  const a = parse(left);
  const b = parse(right);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if ((a[i] || 0) > (b[i] || 0)) return 1;
    if ((a[i] || 0) < (b[i] || 0)) return -1;
  }
  return 0;
}


function validateReleaseForAutomaticUpdate_(metadata, payload) {
  if (!metadata || metadata.channel !== "stable" && metadata.channel !== "development") {
    throw new Error("Release metadata has an unsupported channel.");
  }
  if (metadata.version !== payload.version || metadata.payloadSha256 !== payload.payloadSha256) {
    throw new Error("Release metadata does not match the embedded payload.");
  }
  if (metadata.updateMode !== "automatic") return { allowed: false, reason: "Manual update required." };
  if (
    metadata.minUpdaterVersion &&
    compareSemanticVersions_(SENDMEBOT_INSTALLER_VERSION, metadata.minUpdaterVersion) < 0
  ) {
    return { allowed: false, reason: "A newer SendMeBot Installer is required." };
  }
  if (metadata.requiresReauthorization) {
    return { allowed: false, reason: "The release requires user reauthorization." };
  }
  return { allowed: true, reason: "" };
}


function updateOneRegistration_(target, metadata) {
  const spreadsheet = SpreadsheetApp.openById(target.spreadsheetId);
  const registry = readSendMeBotRegistry_(spreadsheet);
  const verified = verifyUpdateTarget_(target.spreadsheetId, registry);
  if (String(registry.scriptId) !== String(target.scriptId)) {
    throw new Error("The stored update registration no longer matches the workbook registry.");
  }

  const payload = getEmbeddedPayload_();
  const catalog = getEmbeddedReleaseCatalog_();
  const release = validateReleaseForAutomaticUpdate_(metadata, payload);
  if (!release.allowed) return { state: "pending", message: release.reason };
  if (!isTargetInRollout_(target.spreadsheetId, metadata.rolloutPercent)) {
    return { state: "deferred", message: "This installation is outside the current rollout." };
  }

  const currentFiles = verified.files;
  const currentHash = hashAppsScriptFiles_(currentFiles);
  if (currentHash === payload.payloadSha256) {
    return { state: "current", version: payload.version, contentHash: currentHash };
  }
  if (!isKnownReleaseHash_(currentHash, catalog)) {
    throw new Error("The target source differs from every known SendMeBot release; automatic update refused.");
  }
  if (hasNewManifestScopes_(currentFiles, payload.files)) {
    return { state: "pending", message: "The release adds OAuth scopes and requires manual authorization." };
  }

  const path = "projects/" + encodeURIComponent(target.scriptId) + "/content";
  let updateWritten = false;
  try {
    requestAppsScriptApi_("put", path, { files: payload.files });
    updateWritten = true;
    const updated = requestAppsScriptApi_("get", path);
    const updatedHash = hashAppsScriptFiles_(updated.files || []);
    if (updatedHash !== payload.payloadSha256) {
      throw new Error("Updated content hash did not match the embedded payload.");
    }
    return { state: "updated", version: payload.version, contentHash: updatedHash };
  } catch (updateErr) {
    if (!updateWritten) throw updateErr;
    try {
      requestAppsScriptApi_("put", path, { files: currentFiles });
    } catch (rollbackErr) {
      throw new Error(
        "Update verification failed and rollback also failed: " + (rollbackErr.message || rollbackErr)
      );
    }
    throw new Error(
      "Update verification failed; the previous source was restored. " +
      (updateErr.message || updateErr)
    );
  }
}


function runAutomaticUpdates() {
  const targets = getUpdateTargets_();
  if (!targets.length) return { checked: 0, updated: 0, failed: 0, results: [] };
  let metadata;
  try {
    metadata = fetchReleaseMetadata_();
  } catch (err) {
    return {
      checked: targets.length,
      updated: 0,
      failed: targets.length,
      results: targets.map(target => ({ spreadsheetId: target.spreadsheetId, error: err.message || String(err) }))
    };
  }

  const summary = { checked: 0, updated: 0, failed: 0, results: [] };
  targets.forEach(target => {
    summary.checked++;
    try {
      const result = updateOneRegistration_(target, metadata);
      if (result.state === "updated") summary.updated++;
      target.lastCheck = new Date().toISOString();
      target.lastResult = result.state;
      target.installedVersion = result.version || target.installedVersion;
      target.lastContentHash = result.contentHash || target.lastContentHash;
      saveUpdateTarget_(target);
      summary.results.push({ spreadsheetId: target.spreadsheetId, result: result });
    } catch (err) {
      summary.failed++;
      target.lastCheck = new Date().toISOString();
      target.lastResult = "Failed: " + (err.message || err);
      saveUpdateTarget_(target);
      summary.results.push({ spreadsheetId: target.spreadsheetId, error: err.message || String(err) });
    }
  });
  return summary;
}
