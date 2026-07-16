function onHomepage(e) {
  return buildInstallerHomeCard_(e);
}


function onFileScopeGranted(e) {
  return buildInstallerHomeCard_(e);
}


function buildInstallerHomeCard_(e) {
  const event = e || {};
  const sheetsContext = getInstallerSheetsContext_(event);
  const hasFileScope = sheetsContext.hasFileScope;
  const ss = getCurrentInstallerSpreadsheet_(event);
  const sourceName = sheetsContext.title || "this spreadsheet";
  const section = CardService.newCardSection();

  if (!hasFileScope) {
    section
      .setHeader("Create a SendMeBot workbook")
      .addWidget(CardService.newTextInput()
        .setFieldName("newWorkbookName")
        .setTitle("Workbook name")
        .setValue("SendMeBot"))
      .addWidget(CardService.newTextButton()
        .setText("Create fresh SendMeBot workbook")
        .setOnClickAction(CardService.newAction().setFunctionName("createNewSendMeBotWorkbook")))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextParagraph().setText(
        "To preserve sheets from this workbook, grant SendMeBot access to it first."
      ))
      .addWidget(CardService.newTextButton()
      .setText("Grant access to this spreadsheet")
      .setOnClickAction(CardService.newAction().setFunctionName("requestCurrentFileScope")));
  } else {
    const migrationParameters = {
      sourceSpreadsheetId: sheetsContext.spreadsheetId,
      sourceSpreadsheetName: sourceName
    };
    section
      .setHeader("Install SendMeBot into " + sourceName)
      .addWidget(CardService.newTextParagraph().setText(
        "SendMeBot will create a new workbook and copy all worksheet tabs. The original will not be changed."
      ))
      .addWidget(CardService.newTextInput()
        .setFieldName("newWorkbookName")
        .setTitle("Workbook name")
        .setValue(sourceName + " — SendMeBot"))
      .addWidget(CardService.newTextButton()
        .setText("Install SendMeBot into " + sourceName)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction()
          .setFunctionName("installEntireCurrentWorkbook")
          .setParameters(migrationParameters)))
      .addWidget(CardService.newTextButton()
        .setText("Choose which sheets to copy")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("showWorkbookMigrationOptions")
          .setParameters(migrationParameters)));
  }

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("SendMeBot Installer"))
    .addSection(section);

  if (ss && readSendMeBotRegistry_(ss)) {
    const enabled = getUpdateTargets_().some(target => target.spreadsheetId === ss.getId());
    const updateSection = CardService.newCardSection()
      .setHeader("Automatic updates")
      .addWidget(CardService.newTextParagraph().setText(
        enabled
          ? "Automatic updates are enabled for this workbook."
          : "Optional updates require Apps Script API access in your Google Account settings."
      ))
      .addWidget(CardService.newTextButton()
        .setText("Open Apps Script API settings")
        .setOpenLink(CardService.newOpenLink().setUrl(
          "https://script.google.com/home/usersettings"
        )))
      .addWidget(CardService.newTextButton()
        .setText(enabled ? "Disable automatic updates" : "Enable automatic updates")
        .setOnClickAction(CardService.newAction().setFunctionName(
          enabled ? "disableAutomaticUpdates" : "enableAutomaticUpdates"
        )));
    builder.addSection(updateSection);
  }

  return builder.build();
}


function buildEntireWorkbookWarningResponse_(e, ss, scan) {
  const destinationName = getInstallerDestinationName_(e, ss);
  const section = CardService.newCardSection()
    .setHeader("Review workbook compatibility")
    .addWidget(CardService.newTextParagraph().setText(
      "SendMeBot will copy all " + scan.sheets.length + " worksheet tab(s) into the installed workbook."
    ))
    .addWidget(CardService.newTextInput()
      .setFieldName("destinationWorkbookName")
      .setTitle("New workbook name")
      .setValue(destinationName))
    .addWidget(CardService.newTextParagraph().setText(
      "Compatibility warnings:\n• " + scan.warnings.join("\n• ")
    ))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("acceptComplexMigration")
      .setTitle("Complex workbook confirmation")
      .addItem("Continue after reviewing these warnings", "accepted", false));

  if (scan.notices && scan.notices.length) {
    section.addWidget(CardService.newTextParagraph().setText(scan.notices.join("\n")));
  }

  section.addWidget(CardService.newTextButton()
    .setText("Install full workbook")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName("installEntireCurrentWorkbook")
      .setParameters({
        sourceSpreadsheetId: ss.getId(),
        sourceSpreadsheetName: ss.getName()
      })));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("SendMeBot Installer"))
    .addSection(section)
    .build();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}


function requestCurrentFileScope() {
  return CardService.newEditorFileScopeActionResponseBuilder()
    .requestFileScopeForActiveDocument()
    .build();
}


function showWorkbookMigrationOptions(e) {
  const ss = getCurrentInstallerSpreadsheet_(e);
  if (!ss) return notifyInstaller_("Open a Google Sheet and try again.");
  const scan = scanWorkbookCompatibility_(ss);
  const selection = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("sheetIds")
    .setTitle("Sheets to copy");

  scan.sheets.forEach(sheet => {
    selection.addItem(sheet.name, String(sheet.sheetId), true);
  });

  const requestedName = normalizeInstallerValue_(getInstallerFormValue_(e, "newWorkbookName"));
  const destinationName = requestedName || ss.getName() + " — SendMeBot";
  const section = CardService.newCardSection()
    .setHeader("Install SendMeBot into " + ss.getName())
    .addWidget(CardService.newTextParagraph().setText(
      "A new SendMeBot workbook will be created. Your original spreadsheet will not be changed."
    ))
    .addWidget(CardService.newTextInput()
      .setFieldName("destinationWorkbookName")
      .setTitle("New workbook name")
      .setValue(destinationName))
    .addWidget(selection);

  if (scan.warnings.length) {
    section.addWidget(CardService.newTextParagraph().setText(
      "Compatibility warnings:\n• " + scan.warnings.join("\n• ")
    ));
    section.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("acceptComplexMigration")
      .setTitle("Complex workbook confirmation")
      .addItem("Continue after reviewing these warnings", "accepted", false));
  }

  if (scan.notices && scan.notices.length) {
    section.addWidget(CardService.newTextParagraph().setText(scan.notices.join("\n")));
  }

  section.addWidget(CardService.newTextButton()
    .setText("Create installed copy")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("installIntoCurrentWorkbook")
      .setParameters({
        sourceSpreadsheetId: ss.getId(),
        sourceSpreadsheetName: ss.getName()
      })));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("SendMeBot Installer"))
    .addSection(section)
    .build();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}


function buildInstallCompleteResponse_(result) {
  const message = result.incomplete
    ? "The installed copy was created with warnings. Review it before sending."
    : "Your SendMeBot workbook is ready.";
  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText(message))
    .addWidget(CardService.newTextParagraph().setText(
      "Open the new workbook to continue the guided tracker, sender, template, and test-email setup."
    ));

  if (result.warnings && result.warnings.length) {
    section.addWidget(CardService.newTextParagraph().setText(
      "Review:\n• " + result.warnings.join("\n• ")
    ));
  }

  section.addWidget(CardService.newTextButton()
    .setText("Open SendMeBot workbook")
    .setOpenLink(CardService.newOpenLink().setUrl(result.url)));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Installation complete"))
    .addSection(section)
    .build();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}


function notifyInstaller_(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}
