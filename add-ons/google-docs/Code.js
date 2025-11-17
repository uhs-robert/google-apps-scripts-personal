// google-docs/Code.js

// TODO: For published add-ons, use createAddonMenu() instead of createMenu()
// See README.md for details on menu-based Editor add-on best practices

/**
 * Adds a custom menu to the Google Docs UI, with custom scripts.
 * This function is triggered on document open.
 * @return {void}
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu("Scripts")
    .addSubMenu(
      ui
        .createMenu("Templates")
        .addItem(
          "Format Meeting Notes",
          "GoogleDocsScripts.formatMeetingNotes",
        ),
    )
    .addSubMenu(
      ui
        .createMenu("Formatting")
        .addItem(
          "Remove Empty Paragraphs",
          "GoogleDocsScripts.removeEmptyParagraphs",
        )
        .addItem(
          "Convert Heading to Title Case",
          "GoogleDocsScripts.convertHeadingsToTitleCase",
        ),
    )
    .addSubMenu(
      ui
        .createMenu("Utilities")
        .addItem(
          "Calculate Cost for Milestones",
          "GoogleDocsScripts.calculateAndUpdateWithDynamicRate",
        ),
    )
    .addToUi();
}
