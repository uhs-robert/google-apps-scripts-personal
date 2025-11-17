// add-ons/google-docs/Code.js

/**
 * Updating this Add On:
 * Deploy a new version after saving changes (Note the version number)
 * Go to App Configuration in Google Workspace Marketplace SDK: https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk?project=uphill-solutions-add-on
 * In the "App Configuration" section, update add-on script version.
 * Click [Save Draft]
 * Go to "Store Listing" tab, click [Publish]
/*

/**
 * Main function to create the add-on menu.
 * This runs when the document is opened or the add-on is installed.
 * @param {object} e - The event object.
 */
function onOpen(_e) {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu("UpHill Solutions Add Ons");
  buildFullMenu(menu);
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
