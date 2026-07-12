// --- Template modal context ---

function getComposeTemplateContext_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getTrackerSheet_();
  const sendersSheet = ss.getSheetByName("Senders");

  const trackerHeaders = getHeaders_(trackerSheet);

  const rowVariables = Object.keys(trackerHeaders)
    .map(toDisplayHeader_)
    .filter(name => normalize_(name) !== "select")
    .sort();

  const senderVariables = [
    "Sender Name",
    "Sender Email",
    "Sender Signature"
  ];

  const imageAssets = sendersSheet
    ? getImageAssetsForComposer_(sendersSheet)
    : [];

  return {
    rowVariables: rowVariables,
    senderVariables: senderVariables,
    imageAssets: imageAssets,
    imageVariables: imageAssets.map(asset => asset.name)
  };
}


function getImageAssetsForComposer_(sendersSheet) {
  const lastRow = sendersSheet.getLastRow();

  if (lastRow < 2) return [];

  // Image table is E:G
  // E = Name, F = Link, G = Width
  const values = sendersSheet.getRange(2, 5, lastRow - 1, 3).getValues();
  const assets = [];

  values.forEach(row => {
    const name = String(row[0] || "").trim();
    const link = String(row[1] || "").trim();
    const width = String(row[2] || "").trim();

    if (!name || !link) return;

    const fileId = extractDriveFileId_(link);

    let dataUrl = "";

    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const contentType = blob.getContentType() || "image/png";
        const base64 = Utilities.base64Encode(blob.getBytes());

        dataUrl = "data:" + contentType + ";base64," + base64;
      } catch (err) {
        Logger.log("IMAGE PREVIEW ERROR for " + name + ": " + err.message);
      }
    }

    assets.push({
      name: name,
      link: link,
      width: width,
      fileId: fileId,
      previewUrl: dataUrl
    });
  });

  return assets.sort((a, b) => a.name.localeCompare(b.name));
}


// --- Template CRUD ---

function saveComposedTemplate(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const headers = getHeaders_(templateSheet);

  const nameCol = headers["name"];
  const subjectCol = headers["subject"];
  const bodyCol = headers["body"];

  if (!nameCol) throw new Error("Missing Name column in Templates sheet.");
  if (!subjectCol) throw new Error("Missing Subject column in Templates sheet.");
  if (!bodyCol) throw new Error("Missing Body column in Templates sheet.");

  const templateName = String(formData.name || "").trim();
  const subject = String(formData.subject || "").trim();
  const body = String(formData.body || "").trim();

  if (!templateName) throw new Error("Template name is required.");
  if (!subject) throw new Error("Subject is required.");
  if (!body) throw new Error("Body is required.");

  const values = templateSheet.getDataRange().getValues();
  const lookupName = normalizeTemplateKey_(templateName);

  let targetRow = null;

  for (let i = 1; i < values.length; i++) {
    const existingName = normalizeTemplateKey_(values[i][nameCol - 1]);

    if (existingName === lookupName) {
      targetRow = i + 1;
      break;
    }
  }

  if (!targetRow) {
    targetRow = templateSheet.getLastRow() + 1;
  }

  templateSheet.getRange(targetRow, nameCol).setValue(templateName);
  templateSheet.getRange(targetRow, subjectCol).setValue(subject);

  const bodyRange = templateSheet.getRange(targetRow, bodyCol);
  bodyRange.setRichTextValue(buildRichTextValueFromTemplateHtml_(body));
  bodyRange.setWrap(true);

  refreshFormCache();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Template saved: " + templateName,
    "SendMeBot",
    5
  );

  return {
    title: "Template saved",
    message: templateName
  };
}


function getTemplateByKey_(sheet, key) {
  const values = sheet.getDataRange().getValues();
  const headers = getHeaders_(sheet);

  const keyCol = getTemplateNameCol_(headers);
  const subjectCol = headers["subject"];
  const bodyCol = headers["body"];
  const attachmentCol = headers["attachment link"];

  if (!keyCol) throw new Error("Missing Name column in Templates sheet.");
  if (!subjectCol) throw new Error("Missing Subject column.");
  if (!bodyCol) throw new Error("Missing Body column.");

  const lookupKey = normalizeTemplateKey_(key);

  for (let i = 1; i < values.length; i++) {
    if (normalizeTemplateKey_(values[i][keyCol - 1]) === lookupKey) {
      const row = i + 1;
      const bodyCell = sheet.getRange(row, bodyCol);
      const bodyRichText = bodyCell.getRichTextValue();
      const bodyFallback = bodyCell.getDisplayValue();

      return {
        subject: values[i][subjectCol - 1],
        body: richTextValueToTemplateHtml_(bodyRichText, bodyFallback),
        attachmentLink: attachmentCol ? values[i][attachmentCol - 1] : ""
      };
    }
  }

  throw new Error("Template not found: " + key);
}


function getTemplateNameCol_(headers) {
  return headers["name"] || headers["template key"];
}


function getTemplateKeys_(templateSheet) {
  const values = templateSheet.getDataRange().getValues();
  const headers = getHeaders_(templateSheet);
  const keyCol = getTemplateNameCol_(headers);

  if (!keyCol) throw new Error("Missing Name column in Templates sheet.");

  const keys = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][keyCol - 1] || "").trim();
    if (key) keys.push(key);
  }

  return keys;
}

function getEditTemplateContext_() {
  const context = getComposeTemplateContext_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  context.templates = getTemplatesForEditor_(templateSheet);

  return context;
}


function getTemplatesForEditor_(templateSheet) {
  const values = templateSheet.getDataRange().getValues();
  const headers = getHeaders_(templateSheet);

  const nameCol = getTemplateNameCol_(headers);
  const subjectCol = headers["subject"];
  const bodyCol = headers["body"];

  if (!nameCol) throw new Error("Missing Name column in Templates sheet.");
  if (!subjectCol) throw new Error("Missing Subject column in Templates sheet.");
  if (!bodyCol) throw new Error("Missing Body column in Templates sheet.");

  const templates = [];

  for (let i = 1; i < values.length; i++) {
    const row = i + 1;
    const name = String(values[i][nameCol - 1] || "").trim();

    if (!name) continue;

    const bodyCell = templateSheet.getRange(row, bodyCol);
    const bodyRichText = bodyCell.getRichTextValue();
    const bodyFallback = bodyCell.getDisplayValue();

    templates.push({
      name: name,
      subject: String(values[i][subjectCol - 1] || ""),
      body: richTextValueToTemplateHtml_(bodyRichText, bodyFallback)
    });
  }

  return templates.sort((a, b) => a.name.localeCompare(b.name));
}


function getTemplateForEditor(templateName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const template = getTemplateByKey_(templateSheet, templateName);

  return {
    name: templateName,
    subject: template.subject || "",
    body: template.body || ""
  };
}


function updateComposedTemplate(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const headers = getHeaders_(templateSheet);

  const nameCol = headers["name"];
  const subjectCol = headers["subject"];
  const bodyCol = headers["body"];

  if (!nameCol) throw new Error("Missing Name column in Templates sheet.");
  if (!subjectCol) throw new Error("Missing Subject column in Templates sheet.");
  if (!bodyCol) throw new Error("Missing Body column in Templates sheet.");

  const originalName = String(formData.originalName || "").trim();
  const subject = String(formData.subject || "").trim();
  const body = String(formData.body || "").trim();

  if (!originalName) throw new Error("Original template name is required.");
  if (!subject) throw new Error("Subject is required.");
  if (!body) throw new Error("Body is required.");

  const values = templateSheet.getDataRange().getValues();
  const lookupName = normalizeTemplateKey_(originalName);

  let targetRow = null;

  for (let i = 1; i < values.length; i++) {
    const existingName = normalizeTemplateKey_(values[i][nameCol - 1]);

    if (existingName === lookupName) {
      targetRow = i + 1;
      break;
    }
  }

  if (!targetRow) {
    throw new Error("Template not found: " + originalName);
  }

  templateSheet.getRange(targetRow, subjectCol).setValue(subject);

  const bodyRange = templateSheet.getRange(targetRow, bodyCol);
  bodyRange.setRichTextValue(buildRichTextValueFromTemplateHtml_(body));
  bodyRange.setWrap(true);

  refreshFormCache();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Template updated: " + originalName,
    "SendMeBot",
    5
  );

  return {
    title: "Template updated",
    message: originalName
  };
}

// --- Tracker template columns ---

function ensureTrackerColumnForTemplate_(templateKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTrackerSheet_();

  const headers = getHeaders_(sheet);
  const header = getTrackerHeaderForTemplate_(templateKey);
  const normalizedHeader = normalize_(header);

  if (headers[normalizedHeader]) return headers[normalizedHeader];

  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(header);

  applyTemplateColumnConditionalFormatting_(sheet, newCol);

  return newCol;
}

function applyTemplateColumnConditionalFormatting_(sheet, col) {
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const dataRange = sheet.getRange(2, col, maxRows - 1, 1);
  const columnLetter = columnToLetter_(col);
  const firstDataCell = "$" + columnLetter + "2";

  const sentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(TO_TEXT(' + firstDataCell + '),"^Sent on")')
    .setBackground("#d9ead3")
    .setRanges([dataRange])
    .build();

  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(TO_TEXT(' + firstDataCell + '),"^Error:")')
    .setBackground("#f4cccc")
    .setRanges([dataRange])
    .build();

  const inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(LOWER(TO_TEXT(' + firstDataCell + ')),"send|schedul")')
    .setBackground("#fff2cc")
    .setRanges([dataRange])
    .build();

  const existingRules = sheet.getConditionalFormatRules();
  const filteredRules = existingRules.filter(rule =>
    !conditionalFormatRuleTouchesColumn_(rule, sheet, col)
  );

  sheet.setConditionalFormatRules(
    filteredRules.concat([
      sentRule,
      errorRule,
      inProgressRule
    ])
  );
}


function conditionalFormatRuleTouchesColumn_(rule, sheet, col) {
  return rule.getRanges().some(range => {
    if (range.getSheet().getSheetId() !== sheet.getSheetId()) {
      return false;
    }

    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;

    return col >= startCol && col <= endCol;
  });
}


function columnToLetter_(column) {
  let temp = "";
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function getTrackerHeaderForTemplate_(templateKey) {
  let base = normalizeTemplateKey_(templateKey)
    .replace(/\semail$/, "")
    .replace(/\s+/g, " ")
    .trim();

  return toTitleCase_(base) + " Email";
}


function getTemplateStatusColumn_(headers, key) {
  const header = getTrackerHeaderForTemplate_(key);
  const normalizedHeader = normalize_(header);

  return headers[normalizedHeader] || null;
}


// --- Template rendering ---

function renderTextWithAssetsToHtml_(text, textVars, imageAssets, inlineImages, richHtmlVars) {
  const runs = parseTemplateHtmlRuns_(String(text || ""));

  return runs.map(run => {
    const parts = run.text.split(/(\{\{.*?\}\})/g);

    return parts.map(part => {
      const match = part.match(/^\{\{(.*?)\}\}$/);

      if (!match) {
        return applyTemplateRunFormattingHtml_(textToHtml_(part), run);
      }

      const key = match[1].trim();

      if (richHtmlVars && Object.prototype.hasOwnProperty.call(richHtmlVars, key)) {
        return richHtmlVars[key] || "";
      }

      if (Object.prototype.hasOwnProperty.call(textVars, key)) {
        return applyTemplateRunFormattingHtml_(textToHtml_(textVars[key]), run);
      }

      const asset = imageAssets[normalizeTemplateKey_(key)];

      if (asset) {
        return getImageHtmlForAsset_(asset, inlineImages);
      }

      return "";
    }).join("");
  }).join("");
}


function renderPlainTextWithAssets_(text, textVars, imageAssets, richPlainVars) {
  const runs = parseTemplateHtmlRuns_(String(text || ""));

  return runs.map(run => {
    const parts = run.text.split(/(\{\{.*?\}\})/g);

    return parts.map(part => {
      const match = part.match(/^\{\{(.*?)\}\}$/);

      if (!match) {
        return part;
      }

      const key = match[1].trim();

      if (richPlainVars && Object.prototype.hasOwnProperty.call(richPlainVars, key)) {
        return richPlainVars[key] || "";
      }

      if (Object.prototype.hasOwnProperty.call(textVars, key)) {
        return String(textVars[key] || "");
      }

      const asset = imageAssets[normalizeTemplateKey_(key)];

      if (asset) {
        return "[Image: " + asset.name + "]";
      }

      return "";
    }).join("");
  }).join("");
}


function templateBodyTextToHtml_(text) {
  const runs = parseTemplateHtmlRuns_(String(text || ""));

  return runs.map(run => {
    return applyTemplateRunFormattingHtml_(textToHtml_(run.text), run);
  }).join("");
}


function htmlTemplateToPlainText_(text) {
  const runs = parseTemplateHtmlRuns_(String(text || ""));

  return runs.map(run => {
    if (run.href) {
      return run.text + " (" + run.href + ")";
    }

    return run.text;
  }).join("");
}


function buildRichTextValueFromTemplateHtml_(templateHtml) {
  const runs = parseTemplateHtmlRuns_(String(templateHtml || ""));

  let plainText = "";
  const formattedRuns = [];

  runs.forEach(run => {
    const start = plainText.length;
    plainText += run.text;
    const end = plainText.length;

    if (start < end) {
      formattedRuns.push({
        start: start,
        end: end,
        href: run.href || "",
        bold: !!run.bold,
        italic: !!run.italic,
        underline: !!run.underline,
        strikethrough: !!run.strikethrough,
        color: run.color || "",
        fontSize: run.fontSize || "",
        fontFamily: run.fontFamily || ""
      });
    }
  });

  const builder = SpreadsheetApp.newRichTextValue().setText(plainText);

  formattedRuns.forEach(run => {
    if (run.href) {
      builder.setLinkUrl(run.start, run.end, run.href);
    }

    const styleBuilder = SpreadsheetApp.newTextStyle();

    let hasStyle = false;

    if (run.bold) {
      styleBuilder.setBold(true);
      hasStyle = true;
    }

    if (run.italic) {
      styleBuilder.setItalic(true);
      hasStyle = true;
    }

    if (run.underline) {
      styleBuilder.setUnderline(true);
      hasStyle = true;
    }

    if (run.strikethrough) {
      styleBuilder.setStrikethrough(true);
      hasStyle = true;
    }

    if (run.color) {
      styleBuilder.setForegroundColor(run.color);
      hasStyle = true;
    }

    if (run.fontSize) {
      styleBuilder.setFontSize(Number(run.fontSize));
      hasStyle = true;
    }

    if (run.fontFamily) {
      styleBuilder.setFontFamily(run.fontFamily);
      hasStyle = true;
    }

    if (hasStyle) {
      builder.setTextStyle(run.start, run.end, styleBuilder.build());
    }
  });

  return builder.build();
}


function richTextValueToTemplateHtml_(richTextValue, fallbackText) {
  if (!richTextValue) {
    return String(fallbackText || "");
  }

  const fullText = richTextValue.getText();

  if (!fullText) {
    return "";
  }

  const runs = richTextValue.getRuns();

  if (!runs || !runs.length) {
    return escapeHtmlPreservingNewlines_(fullText);
  }

  return runs.map(run => {
    const style = run.getTextStyle();

    const templateRun = {
      text: run.getText(),
      href: run.getLinkUrl() || "",
      bold: style ? style.isBold() : false,
      italic: style ? style.isItalic() : false,
      underline: style ? style.isUnderline() : false,
      strikethrough: style ? style.isStrikethrough() : false,
      color: style ? normalizeCssColor_(style.getForegroundColor()) : "",
      fontSize: style && shouldPreserveRichTextFontSize_(style.getFontSize())
        ? style.getFontSize()
        : "",
      fontFamily: style ? style.getFontFamily() : ""
    };

    return applyTemplateRunFormattingHtml_(
      escapeHtmlPreservingNewlines_(templateRun.text),
      templateRun
    );
  }).join("");
}


function escapeHtmlPreservingNewlines_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}


function stripHtmlTags_(value) {
  return String(value || "").replace(/<[^>]*>/g, "");
}


function decodeHtmlEntities_(value) {
  return String(value || "")
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function parseTemplateHtmlRuns_(html) {
  const raw = String(html || "");
  const tagPattern = /<\/?(a|strong|b|em|i|u|s|strike|span|font)\b[^>]*>/gi;

  const runs = [];
  let lastIndex = 0;
  let match;

  const state = {
    hrefStack: [],
    boldDepth: 0,
    italicDepth: 0,
    underlineDepth: 0,
    strikethroughDepth: 0,
    styleStack: []
  };

  while ((match = tagPattern.exec(raw)) !== null) {
    addTemplateHtmlTextRun_(
      runs,
      raw.slice(lastIndex, match.index),
      state
    );

    const tag = match[0];
    const tagName = match[1].toLowerCase();
    const isClosing = /^<\//.test(tag);

    if (isClosing) {
      closeTemplateTag_(state, tagName);
    } else {
      openTemplateTag_(state, tagName, tag);
    }

    lastIndex = tagPattern.lastIndex;
  }

  addTemplateHtmlTextRun_(
    runs,
    raw.slice(lastIndex),
    state
  );

  return mergeAdjacentTemplateRuns_(runs);
}


function openTemplateTag_(state, tagName, tag) {
  if (tagName === "a") {
    const href = decodeHtmlEntities_(getHtmlAttribute_(tag, "href"));

    if (/^(https?:\/\/|mailto:)/i.test(href)) {
      state.hrefStack.push(href);
    } else {
      state.hrefStack.push("");
    }

    return;
  }

  if (tagName === "strong" || tagName === "b") {
    state.boldDepth++;
    return;
  }

  if (tagName === "em" || tagName === "i") {
    state.italicDepth++;
    return;
  }

  if (tagName === "u") {
    state.underlineDepth++;
    return;
  }

  if (tagName === "s" || tagName === "strike") {
    state.strikethroughDepth++;
    return;
  }

  if (tagName === "span") {
    state.styleStack.push(parseStyleAttribute_(getHtmlAttribute_(tag, "style")));
    return;
  }

  if (tagName === "font") {
    state.styleStack.push({
      color: normalizeCssColor_(getHtmlAttribute_(tag, "color")),
      fontFamily: sanitizeFontFamily_(getHtmlAttribute_(tag, "face")),
      fontSize: ""
    });
  }
}


function closeTemplateTag_(state, tagName) {
  if (tagName === "a") {
    state.hrefStack.pop();
    return;
  }

  if (tagName === "strong" || tagName === "b") {
    state.boldDepth = Math.max(0, state.boldDepth - 1);
    return;
  }

  if (tagName === "em" || tagName === "i") {
    state.italicDepth = Math.max(0, state.italicDepth - 1);
    return;
  }

  if (tagName === "u") {
    state.underlineDepth = Math.max(0, state.underlineDepth - 1);
    return;
  }

  if (tagName === "s" || tagName === "strike") {
    state.strikethroughDepth = Math.max(0, state.strikethroughDepth - 1);
    return;
  }

  if (tagName === "span" || tagName === "font") {
    state.styleStack.pop();
  }
}


function addTemplateHtmlTextRun_(runs, text, state) {
  if (!text) return;

  const currentStyle = getCurrentTemplateStyle_(state);

  runs.push({
    text: decodeHtmlEntities_(stripHtmlTags_(text)),
    href: currentStyle.href,
    bold: currentStyle.bold,
    italic: currentStyle.italic,
    underline: currentStyle.underline,
    strikethrough: currentStyle.strikethrough,
    color: currentStyle.color,
    fontSize: currentStyle.fontSize,
    fontFamily: currentStyle.fontFamily
  });
}


function getCurrentTemplateStyle_(state) {
  const style = {
    href: state.hrefStack.length ? state.hrefStack[state.hrefStack.length - 1] : "",
    bold: state.boldDepth > 0,
    italic: state.italicDepth > 0,
    underline: state.underlineDepth > 0,
    strikethrough: state.strikethroughDepth > 0,
    color: "",
    fontSize: "",
    fontFamily: ""
  };

  state.styleStack.forEach(styleLayer => {
    if (styleLayer.color) style.color = styleLayer.color;
    if (styleLayer.fontSize) style.fontSize = styleLayer.fontSize;
    if (styleLayer.fontFamily) style.fontFamily = styleLayer.fontFamily;
  });

  return style;
}


function mergeAdjacentTemplateRuns_(runs) {
  const merged = [];

  runs.forEach(run => {
    if (!run.text) return;

    const previous = merged[merged.length - 1];

    if (
      previous &&
      previous.href === run.href &&
      previous.bold === run.bold &&
      previous.italic === run.italic &&
      previous.underline === run.underline &&
      previous.strikethrough === run.strikethrough &&
      previous.color === run.color &&
      previous.fontSize === run.fontSize &&
      previous.fontFamily === run.fontFamily
    ) {
      previous.text += run.text;
    } else {
      merged.push(run);
    }
  });

  return merged;
}


function applyTemplateRunFormattingHtml_(htmlText, run) {
  if (!htmlText) return "";

  let output = htmlText;
  let css = "";

  if (run.color) {
    css += "color:" + run.color + ";";
  }

  if (run.fontSize) {
    css += "font-size:" + Number(run.fontSize) + "pt;";
  }

  if (run.fontFamily) {
    css += "font-family:" + escapeHtml_(run.fontFamily) + ", Arial, sans-serif;";
  }

  if (css) {
    output = '<span style="' + css + '">' + output + '</span>';
  }

  if (run.strikethrough) {
    output = "<s>" + output + "</s>";
  }

  if (run.underline) {
    output = "<u>" + output + "</u>";
  }

  if (run.italic) {
    output = "<em>" + output + "</em>";
  }

  if (run.bold) {
    output = "<strong>" + output + "</strong>";
  }

  if (run.href && /^(https?:\/\/|mailto:)/i.test(run.href)) {
    output =
      '<a href="' +
      escapeHtml_(run.href) +
      '" target="_blank">' +
      output +
      '</a>';
  }

  return output;
}


function parseStyleAttribute_(style) {
  const raw = String(style || "");
  const result = {
    color: "",
    fontSize: "",
    fontFamily: ""
  };

  raw.split(";").forEach(rule => {
    const parts = rule.split(":");
    if (parts.length < 2) return;

    const property = parts[0].trim().toLowerCase();
    const value = parts.slice(1).join(":").trim();

    if (property === "color") {
      result.color = normalizeCssColor_(value);
    }

    if (property === "font-size") {
      result.fontSize = sanitizeFontSize_(value);
    }

    if (property === "font-family") {
      result.fontFamily = sanitizeFontFamily_(value);
    }
  });

  return result;
}


function getHtmlAttribute_(tag, attrName) {
  const pattern = new RegExp(attrName + '\\s*=\\s*["\\\']([^"\\\']*)["\\\']', "i");
  const match = String(tag || "").match(pattern);

  return match && match[1] ? match[1] : "";
}


function normalizeCssColor_(value) {
  const raw = String(value || "").trim();

  if (!raw) return "";

  if (/^#[0-9a-f]{3}$/i.test(raw) || /^#[0-9a-f]{6}$/i.test(raw)) {
    return raw.toLowerCase();
  }

  const rgbMatch = raw.match(/^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i);

  if (rgbMatch) {
    const r = Math.max(0, Math.min(255, Number(rgbMatch[1])));
    const g = Math.max(0, Math.min(255, Number(rgbMatch[2])));
    const b = Math.max(0, Math.min(255, Number(rgbMatch[3])));

    return "#" +
      [r, g, b]
        .map(n => n.toString(16).padStart(2, "0"))
        .join("");
  }

  return "";
}


function sanitizeFontSize_(value) {
  const raw = String(value || "").trim().toLowerCase();

  if (!raw) return "";

  const match = raw.match(/^(\d{1,3})(px|pt)?$/);

  if (!match) return "";

  const size = Number(match[1]);

  if (!size || size < 6 || size > 72) return "";

  return String(size);
}


function sanitizeFontFamily_(value) {
  const raw = String(value || "")
    .replace(/["']/g, "")
    .trim();

  if (!raw) return "";

  const firstFamily = raw.split(",")[0].trim();

  if (!/^[a-zA-Z0-9\s_-]+$/.test(firstFamily)) {
    return "";
  }

  return firstFamily;
}

function getTemplateFormContext_(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getTrackerSheet_();
  const sendersSheet = ss.getSheetByName("Senders");
  const templateSheet = ss.getSheetByName("Templates");

  if (!templateSheet) throw new Error("Missing Templates sheet.");

  const trackerHeaders = getHeaders_(trackerSheet);

  const rowVariables = Object.keys(trackerHeaders)
    .map(toDisplayHeader_)
    .filter(name => normalize_(name) !== "select")
    .sort();

  const senderVariables = [
    "Sender Name",
    "Sender Email",
    "Sender Signature"
  ];

  const imageAssets = sendersSheet
    ? getImageAssetsForComposer_(sendersSheet)
    : [];

  const templates = mode === "edit"
    ? getTemplateKeys_(templateSheet)
    : [];

  return {
    mode: mode || "new",
    templates: templates,
    rowVariables: rowVariables,
    senderVariables: senderVariables,
    imageAssets: imageAssets,
    imageVariables: imageAssets.map(asset => asset.name)
  };
}
