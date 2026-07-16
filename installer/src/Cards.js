function onHomepage(e) {
  return buildInstallerHomeCard_(e);
}


function onFileScopeGranted(e) {
  return buildInstallerHomeCard_(e);
}


function buildInstallerHomeCard_(e) {
  const event = e || {};
  const hasFileScope = !(
    event.sheets && event.sheets.addonHasFileScopePermission === false
  );
  const ss = hasFileScope ? getCurrentInstallerSpreadsheet_() : null;
  const sourceName = ss ? ss.getName() : "this spreadsheet";
  const section = CardService.newCardSection()
    .setHeader("Create a SendMeBot workbook")
    .addWidget(CardService.newTextInput()
      .setFieldName("newWorkbookName")
      .setTitle("Workbook name")
      .setValue("SendMeBot"))
    .addWidget(CardService.newTextButton()
      .setText("Create new SendMeBot spreadsheet")
      .setOnClickAction(CardService.newAction().setFunctionName("createNewSendMeBotWorkbook")));

  if (!hasFileScope) {
    section.addWidget(CardService.newDivider());
    section.addWidget(CardService.newTextButton()
      .setText("Grant access to this spreadsheet")
      .setOnClickAction(CardService.newAction().setFunctionName("requestCurrentFileScope")));
  } else if (ss) {
    section.addWidget(CardService.newDivider());
    section.addWidget(CardService.newTextButton()
      .setText("Install SendMeBot into " + sourceName)
      .setOnClickAction(CardService.newAction().setFunctionName("showWorkbookMigrationOptions")));
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


function requestCurrentFileScope() {
  return CardService.newEditorFileScopeActionResponseBuilder()
    .requestFileScopeForActiveDocument()
    .build();
}


function showWorkbookMigrationOptions() {
  const ss = getCurrentInstallerSpreadsheet_();
  if (!ss) return notifyInstaller_("Open a Google Sheet and try again.");
  const scan = scanWorkbookCompatibility_(ss);
  const selection = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("sheetIds")
    .setTitle("Sheets to copy");

  scan.sheets.forEach(sheet => {
    selection.addItem(sheet.name, String(sheet.sheetId), true);
  });

  const section = CardService.newCardSection()
    .setHeader("Install SendMeBot into " + ss.getName())
    .addWidget(CardService.newTextParagraph().setText(
      "A new SendMeBot workbook will be created. Your original spreadsheet will not be changed."
    ))
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
    .setOnClickAction(CardService.newAction().setFunctionName("installIntoCurrentWorkbook")));

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
      "Open the new workbook, then use SendMeBot → Setup to choose the tracker and Record ID column."
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
