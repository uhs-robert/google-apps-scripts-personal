// add-ons/google-docs/Code.js

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
