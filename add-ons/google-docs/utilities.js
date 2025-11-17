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

  if (section.addHR) {
    body.insertHorizontalRule(index);
  }

  paragraph.setFontFamily(section.font).setHeading(section.heading);
  return true;
}

/**
 * Deletes a section and optionally the paragraph underneath it.
 * @param {GoogleAppsScript.Document.Body} body - The document body containing the section
 * @param {string} sectionName - The section name to search for and delete
 * @param {boolean} [deleteNextParagraph=false] - Whether to delete the paragraph underneath the section
 * @return {boolean} True if the section was found and deleted, false otherwise
 */
function deleteSection(body, sectionName, deleteNextParagraph = false) {
  const foundText = body.findText(sectionName);
  if (!foundText) return false;

  const paragraph = foundText.getElement().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);

  paragraph.removeFromParent();

  if (deleteNextParagraph) {
    const numChildren = body.getNumChildren();
    if (index < numChildren) {
      const nextElement = body.getChild(index);
      if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
        nextElement.removeFromParent();
      }
    }
  }

  return true;
}
