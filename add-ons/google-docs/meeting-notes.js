// add-ons/google-docs/meeting-notes.js

/**
 * Finds the meeting date in the document body or defaults to today's date.
 * @param {GoogleAppsScript.Document.Body} body - The document body to search
 * @return {string} The meeting date in YYYY-MM-DD format
 */
function getMeetingDate(body) {
  const datePattern = /\b([A-Za-z]+ \d{1,2})\b/;
  const dateMatch = body.findText(datePattern);

  let dateToFormat;
  if (dateMatch) {
    const matchedText = dateMatch.getElement().asText().getText();
    const result = datePattern.exec(matchedText);
    if (result) {
      const currentYear = new Date().getFullYear();
      dateToFormat = new Date(`${result[0]}, ${currentYear}`);
    }
  }

  if (!dateToFormat) dateToFormat = new Date();

  const year = dateToFormat.getFullYear();
  const month = String(dateToFormat.getMonth() + 1).padStart(2, "0");
  const day = String(dateToFormat.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
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

  // Remove the section heading
  paragraph.removeFromParent();

  // Optionally remove the paragraph underneath
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

/**
 * Formats specific sections, updates {TITLE} in the header with "Meeting Notes - Date",
 * applies custom fonts and headings to the "Notes" section, and deletes the "Action items" section.
 * @return {void}
 */
function formatMeetingNotes() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  const meetingDate = getMeetingDate(body);
  const titleTextToInsert = `${meetingDate} - Meeting Notes`;
  doc.setName(titleTextToInsert);
  updateDocumentHeader(doc, titleTextToInsert);

  const sections = [
    {
      name: "Notes",
      font: "Montserrat",
      heading: DocumentApp.ParagraphHeading.HEADING3,
      addHR: true,
    },
  ];
  sections.forEach((section) => formatSection(body, section));

  // Delete Action items section and the paragraph underneath
  deleteSection(body, "Action items", true);
}
