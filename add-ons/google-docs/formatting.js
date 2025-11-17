// add-ons/google-docs/formatting.js

/**
 * Removes empty paragraphs from the document, preserving list items, checkboxes,
 * bullets, numbered lists, and horizontal rules.
 *
 * @return {void}
 */
function removeEmptyParagraphs() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();
  paragraphs.splice(-1, 1); // remove last paragraph to prevent errors

  paragraphs.forEach((paragraph) => {
    // Check if the paragraph is empty and does not contain other elements like lists or horizontal rules
    const text = paragraph.getText().trim();
    const isListItem =
      paragraph.getType() === DocumentApp.ElementType.LIST_ITEM;
    const hasHorizontalRule =
      paragraph.findElement(DocumentApp.ElementType.HORIZONTAL_RULE) !== null;

    if (!text && !isListItem && !hasHorizontalRule) {
      body.removeChild(paragraph);
    }
  });
}

/**
 * Converts all headings in the document to title case.
 * Only affects paragraphs with a heading style.
 */
function convertHeadingsToTitleCase() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  paragraphs.forEach((paragraph) => {
    // Apply title case only to paragraphs that are headings
    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
      const text = paragraph.getText();
      const titleCasedText = text.replace(
        /\w\S*/g,
        (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase(),
      );
      paragraph.setText(titleCasedText);
    }
  });
}
