// add-ons/google-docs/templates.js

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

/**
 * Main function to set rate, calculate costs, and update totals in the document.
 */
function calculateAndUpdateWithDynamicRate() {
  const response = getUserInput("Enter hourly rate:"); // Step 1: Prompt user for rate
  const rate = parseFloat(response);
  if (isNaN(rate)) {
    Logger.log("Rate not provided. Exiting function.");
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  Logger.log("Starting calculation with rate: $" + rate + "/hour");

  // Step 2: Find the milestone table
  const table = findMilestoneTable(body);
  if (!table) {
    Logger.log("Milestone table not found.");
    return;
  }

  // Step 3: Calculate costs and totals from the milestone table
  const totals = calculateCostsAndTotalsFromTable(table, rate);

  // Step 4: Update totals in the last row of the table
  updateTotalRow(table, totals);
}

/**
 * Finds the table containing the milestone data based on keywords.
 * @param {GoogleAppsScript.Document.Body} body - The document body.
 * @return {GoogleAppsScript.Document.Table|null} - The identified table or null if not found.
 */
function findMilestoneTable(body) {
  const tables = body.getTables();

  Logger.log("Searching through tables...");
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    Logger.log(`Checking table ${i + 1}: ${table.getText()}`);

    // Loosened condition to search for main keywords individually
    if (
      table.getText().includes("Milestone") &&
      table.getText().includes("Est. Time") &&
      table.getText().includes("Est. Cost")
    ) {
      Logger.log(`Milestone table found at index ${i + 1}`);
      return table;
    }
  }
  Logger.log("Milestone table not found in document.");
  return null;
}

/**
 * Calculates the estimated cost and totals from the milestone table based on the rate.
 * @param {GoogleAppsScript.Document.Table} table - The table with milestone data.
 * @param {number} rate - The hourly rate for calculating estimated costs.
 * @return {Object} - An object with totalTime and totalCost.
 */
function calculateCostsAndTotalsFromTable(table, rate) {
  let totalTime = 0;
  let totalCost = 0;

  // Assume the table structure:
  // Row 1: Headers (e.g., "Milestone", "Description", "Est. Time (hours)", "Est. Cost")
  // Subsequent rows: Milestone data

  for (let i = 1; i < table.getNumRows() - 1; i++) {
    // Skip header row and total row
    const row = table.getRow(i);

    // Extract estimated time from the third column (0-based) and produce estimate
    const estTimeText = row.getCell(2).getText();
    const estTime = parseFloat(estTimeText.replace(/[^\d.]/g, "")) || 0;
    const estCost = estTime * rate;

    // Format cost with commas and set it in the document (4th column)
    row
      .getCell(3)
      .setText(
        `$${estCost.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`,
      );

    // Accumulate total time and total cost
    totalTime += estTime;
    totalCost += estCost;
  }

  return {
    totalTime: totalTime,
    totalCost: totalCost,
  };
}

/**
 * Updates the last row of the table with the calculated total time and cost.
 * @param {GoogleAppsScript.Document.Table} table - The table with milestone data.
 * @param {Object} totals - The totals to be displayed, containing totalTime and totalCost.
 */
function updateTotalRow(table, totals) {
  const lastRow = table.getRow(table.getNumRows() - 1); // Last row for totals
  lastRow.getCell(2).setText(`${totals.totalTime} hour(s)`); // Update total time cell
  lastRow
    .getCell(3)
    .setText(
      `$${totals.totalCost.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`,
    );
}
