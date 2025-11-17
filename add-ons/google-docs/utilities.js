// add-ons/google-docs/utilities.js

/**
 * Prompts the user to enter a value.
 * @return {string} - The input provided or null.
 */
function sendUserPrompt(message) {
  const ui = DocumentApp.getUi();
  const response = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    return response.getResponseText();
  }
  return null;
}
