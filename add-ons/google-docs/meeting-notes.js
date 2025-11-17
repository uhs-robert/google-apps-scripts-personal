// add-ons/google-docs/meeting-notes.js

/**
 * Formats specific sections, updates {TITLE} in the header with "Meeting Notes - Date",
 * and applies custom fonts and headings to sections like "Notes" and "Action items."
 * @return {void}
 */
function formatMeetingNotes() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Locate the meeting date in "MMM DD, YYYY" format or default to today
  const datePattern = /\b([A-Za-z]{3} \d{1,2}, \d{4})\b/;
  const dateMatch = body.findText(datePattern);
  let meetingDate;

  if (dateMatch) {
    const matchedText = dateMatch.getElement().asText().getText();
    const result = datePattern.exec(matchedText);
    meetingDate = result ? result[0] : null;
  }

  if (!meetingDate) {
    const currentDate = new Date();
    const year = currentDate.getFullYear();
    const month = String(currentDate.getMonth() + 1).padStart(2, "0");
    const day = String(currentDate.getDate()).padStart(2, "0");
    meetingDate = `${year}-${month}-${day}`;
  }

  // Set the document's name
  const titleTextToInsert = `${meetingDate} - Meeting Notes`;
  doc.setName(titleTextToInsert);

  // Update header {TITLE} placeholder with formatted title
  const header = doc.getHeader();
  if (header) {
    const titleText = header.findText("\\{TITLE\\}");
    if (titleText) {
      const titleElement = titleText.getElement().asText();
      titleElement.setText(titleTextToInsert).setFontSize(10);
    }
  }

  // Define and format sections
  const sections = [
    {
      name: "Notes",
      font: "Montserrat",
      heading: DocumentApp.ParagraphHeading.HEADING2,
      addHR: true,
    },
    {
      name: "Action items",
      font: "Montserrat",
      heading: DocumentApp.ParagraphHeading.HEADING2,
    },
  ];

  sections.forEach((section) => {
    const foundText = body.findText(section.name);
    if (foundText) {
      const paragraph = foundText.getElement().getParent().asParagraph();
      const index = body.getChildIndex(paragraph);
      if (section.addHR) body.insertHorizontalRule(index);
      paragraph.setFontFamily(section.font).setHeading(section.heading);
    }
  });
}
