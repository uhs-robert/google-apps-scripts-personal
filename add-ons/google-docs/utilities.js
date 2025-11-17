// add-ons/google-docs/utilities.js

/**
 * Prompts the user to enter a value.
 * @return {string} - The input provided or null.
 */
function sendUserPrompt(message) {
  const ui = DocumentApp.getUi();
  const response = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK)
    return response.getResponseText();

  return null;
}

/**
 * Updates the document header by replacing {TITLE} placeholder with the provided title.
 * @param {GoogleAppsScript.Document.Document} doc - The document to update
 * @param {string} title - The title text to insert
 * @return {boolean} True if the header was updated, false otherwise
 */
function updateDocumentHeader(doc, title) {
  const header = doc.getHeader();
  if (!header) return false;

  const titleText = header.findText("\\{TITLE\\}");
  if (!titleText) return false;

  const titleElement = titleText.getElement().asText();
  titleElement.setText(title);
  return true;
}

/**
 * Formats a specific section in the document body with custom styling.
 * @param {GoogleAppsScript.Document.Body} body - The document body containing the section
 * @param {Object} section - Section configuration object
 * @param {string} section.name - The section name to search for
 * @param {string} section.font - The font family to apply
 * @param {GoogleAppsScript.Document.ParagraphHeading} section.heading - The heading level
 * @param {boolean} [section.addHR=false] - Whether to add a horizontal rule before the section
 * @return {boolean} True if the section was found and formatted, false otherwise
 */
function formatSection(body, section) {
  const foundText = body.findText(section.name);
  if (!foundText) return false;

  const paragraph = foundText.getElement().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);

  if (section.addHR) body.insertHorizontalRule(index);

  paragraph.setFontFamily(section.font).setHeading(section.heading);
  return true;
}

/**
 * Deletes a section and optionally N elements underneath it.
 * @param {GoogleAppsScript.Document.Body} body - The document body containing the section
 * @param {string} sectionName - The section name to search for and delete
 * @param {number} [deleteNextElements=0] - Number of elements to delete after the section (0 = none)
 * @return {boolean} True if the section was found and deleted, false otherwise
 */
function deleteSection(body, sectionName, deleteNextElements = 0) {
  const foundText = body.findText(sectionName);
  if (!foundText) return false;

  const paragraph = foundText.getElement().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);

  paragraph.removeFromParent();

  for (let i = 0; i < deleteNextElements; i++) {
    const numChildren = body.getNumChildren();
    if (index >= numChildren) break;

    const nextElement = body.getChild(index);
    nextElement.removeFromParent();
  }

  return true;
}

/**
 * Inserts text before the found search text.
 * @param {GoogleAppsScript.Document.Body} body - The document body to search in
 * @param {string} searchText - The text to search for
 * @param {string} textToInsert - The text to insert before the found text
 * @return {GoogleAppsScript.Document.Paragraph|null} The inserted paragraph, or null if search text not found
 */
function insertTextBefore(body, searchText, textToInsert) {
  const foundText = body.findText(searchText);
  if (!foundText) return null;

  const paragraph = foundText.getElement().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);

  return body.insertParagraph(index, textToInsert);
}
