const SENDMEBOT_APP_ID = "sendmebot";
const SENDMEBOT_APP_VERSION = "1.1.0";
const SENDMEBOT_INSTALL_FORMAT_VERSION = 1;
const SENDMEBOT_ENVIRONMENT_CONFIG_PROPERTY = "SENDMEBOT_ENVIRONMENT_CONFIG";

const SENDMEBOT_NEUTRAL_CONFIG = {
  brandName: "SendMeBot",
  primaryColor: "#356854",
  accentColor: "#1a73e8",
  darkColor: "#20342d",
  charcoalColor: "#666666",
  logoUrl: "https://www.gstatic.com/images/icons/material/system/2x/send_black_48dp.png",
  assistantUrl: "https://gemini.google.com/gem/6c4e76611b3f",
  defaultTrackerSheetName: "Tracker",
  defaultRecordIdHeader: "Name"
};


function getSendMeBotEnvironmentConfig_(runtime) {
  const options = runtime || {};
  const properties = options.scriptProperties || (
    typeof PropertiesService !== "undefined" &&
    typeof PropertiesService.getScriptProperties === "function"
      ? PropertiesService.getScriptProperties()
      : null
  );
  let configured = {};

  if (properties && typeof properties.getProperty === "function") {
    const raw = properties.getProperty(SENDMEBOT_ENVIRONMENT_CONFIG_PROPERTY);
    if (raw) {
      try {
        configured = JSON.parse(raw) || {};
      } catch (err) {
        throw new Error("SendMeBot environment configuration is not valid JSON.");
      }
    }
  }

  return sanitizeSendMeBotEnvironmentConfig_(Object.assign(
    {},
    SENDMEBOT_NEUTRAL_CONFIG,
    configured,
    options.overrides || {}
  ));
}


function sanitizeSendMeBotEnvironmentConfig_(value) {
  const config = value || {};
  const color = function(candidate, fallback) {
    const normalized = String(candidate || "").trim();
    return /^#[0-9a-f]{6}$/i.test(normalized) ? normalized : fallback;
  };
  const url = function(candidate, fallback) {
    const normalized = String(candidate || "").trim();
    return /^https:\/\//i.test(normalized) ? normalized : fallback;
  };

  return {
    brandName: String(config.brandName || SENDMEBOT_NEUTRAL_CONFIG.brandName).trim(),
    primaryColor: color(config.primaryColor, SENDMEBOT_NEUTRAL_CONFIG.primaryColor),
    accentColor: color(config.accentColor, SENDMEBOT_NEUTRAL_CONFIG.accentColor),
    darkColor: color(config.darkColor, SENDMEBOT_NEUTRAL_CONFIG.darkColor),
    charcoalColor: color(config.charcoalColor, SENDMEBOT_NEUTRAL_CONFIG.charcoalColor),
    logoUrl: url(config.logoUrl, SENDMEBOT_NEUTRAL_CONFIG.logoUrl),
    assistantUrl: url(config.assistantUrl, SENDMEBOT_NEUTRAL_CONFIG.assistantUrl),
    defaultTrackerSheetName: String(
      config.defaultTrackerSheetName || SENDMEBOT_NEUTRAL_CONFIG.defaultTrackerSheetName
    ).trim(),
    defaultRecordIdHeader: String(
      config.defaultRecordIdHeader || SENDMEBOT_NEUTRAL_CONFIG.defaultRecordIdHeader
    ).trim()
  };
}


function getSendMeBotClientBrand_(runtime) {
  const config = getSendMeBotEnvironmentConfig_(runtime);
  return {
    name: config.brandName,
    primaryColor: config.primaryColor,
    accentColor: config.accentColor,
    darkColor: config.darkColor,
    charcoalColor: config.charcoalColor,
    logoUrl: config.logoUrl,
    assistantUrl: config.assistantUrl
  };
}


function getSendMeBotAppVersion_() {
  return SENDMEBOT_APP_VERSION;
}
