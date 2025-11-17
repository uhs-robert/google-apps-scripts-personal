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

  insertTextBefore(body, "Notes", "@ai");

  const sections = [
    {
      name: "Notes",
      font: "Montserrat",
      heading: DocumentApp.ParagraphHeading.HEADING3,
      addHR: true,
    },
  ];
  sections.forEach((section) => formatSection(body, section));

  deleteSection(body, "Action items", 2);
}
