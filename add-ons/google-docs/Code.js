// add-ons/google-docs/Code.js

// TODO: For published add-ons, use createAddonMenu() instead of createMenu()
// See README.md for details on menu-based Editor add-on best practices

/**
 * Main function to create the add-on menu.
 * This runs when the document is opened or the add-on is installed.
 * @param {object} e - The event object.
 */
function onOpen(e) {
  const ui = DocumentApp.getUi();
  const menu = ui.createAddonMenu("UpHill Solutions Add Ons");

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    menu.addItem("Start & Authorize", "showAuthPrompt");
  } else {
    buildFullMenu(menu);
  }
  menu.addToUi();
}

/**
 * Creates the full menu items for the add-on.
 * @param {GoogleAppsScript.Base.Menu} menu - The parent menu object.
 */
function buildFullMenu(menu) {
  const ui = DocumentApp.getUi();

  menu
    .addSubMenu(
      ui
        .createMenu("Templates")
        .addItem("Format Meeting Notes", "formatMeetingNotes"),
    )
    .addSubMenu(
      ui
        .createMenu("Formatting")
        .addItem("Remove Empty Paragraphs", "removeEmptyParagraphs")
        .addItem("Convert Heading to Title Case", "convertHeadingsToTitleCase"),
    )
    .addSubMenu(
      ui
        .createMenu("Utilities")
        .addItem(
          "Calculate Cost for Milestones",
          "calculateAndUpdateWithDynamicRate",
        ),
    );
}

/**
 * Shows a simple prompt to trigger the authorization flow.
 */
function showAuthPrompt() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    "Authorization Required",
    "This add-on needs your permission to run. Please click OK to authorize, then run the 'Start' menu again.",
    ui.ButtonSet.OK,
  );

  try {
    DocumentApp.getActiveDocument();
  } catch (e) {
    ui.alert(
      "Authorization may not have completed. Please try running 'Start & Authorize' again.",
    );
  }
}
