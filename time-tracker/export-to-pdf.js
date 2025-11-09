// time-tracker/export-to-pdf.js

// === Main ================================================
/**
 * Main orchestrator for PDF export pipeline
 * Handles full workflow from data filtering to PDF generation with user feedback
 * @param {string} companyName - Target company for filtering time entries
 * @param {string} projectName - Target project for filtering time entries
 * @param {boolean} log - Enable detailed logging for debugging
 */
const exportToPDF = (companyName, projectName, log = false) => {
  try {
    const destFolder = DriveApp.getFolderById(destFolderId);

    // Check if the company folder exists, if not, create it
    const companyFolder = getOrCreateFolder(destFolder, companyName);

    if (log) Logger.log("Filtering Data...");
    const data = getFilteredData(companyName, projectName);
    if (data.length === 0)
      throw new Error("No data found for the specified company and project.");

    if (log) Logger.log("Generating File Name...");
    const Dates = getDates(data);
    const fileName = generateFileName(Dates, companyName);

    if (log) Logger.log("Creating Document from Template...");
    const { docFile, docId } = createDocumentFromTemplate(
      templateDocId,
      fileName,
      companyFolder,
    );

    if (log) Logger.log("Updating Doc Placeholders...");
    updateDocumentPlaceholders(docFile, companyName, data, Dates);

    if (log) Logger.log("Converting Document to PDF...");
    const pdfFile = convertDocToPDF(docId, fileName, companyFolder);

    if (log) Logger.log("Showing Success Dialog...");
    const message = `
      <p><strong>${companyName}</strong></p>
      <p>${Dates.formattedStart} to ${Dates.formattedEnd}</p>
      <hr />
      <ol>
        <li><a href="${pdfFile.getUrl()}" target="_blank">Open as PDF</a></li>
        <li><a href="https://docs.google.com/document/d/${docFile.getId()}/edit" target="_blank">Open as Google Doc</a></li>
      </ol>
    `;
    showDialog("PDF Ready", message);
  } catch (error) {
    if (log) Logger.log("Showing Error Dialog...");
    const message = `<p><strong style="color: red;">Error</strong>: ${error.message}</p>`;
    showDialog("Error in PDF Generation", message);
    throw error;
  }
};

/**
 * Populates document template with time tracking data and metadata
 * Handles both text placeholders and dynamic table generation
 * @param {GoogleAppsScript.Document.Document} docFile - Template document to populate
 * @param {string} companyName - Company name for calculations
 * @param {Array[]} data - Time tracking data rows
 * @param {Object} Dates - Date formatting information
 */
const updateDocumentPlaceholders = (docFile, companyName, data, Dates) => {
  const body = docFile.getBody();
  const header = docFile.getHeader();
  const totalHours = data.reduce(
    (sum, row) => sum + row[ColIndex.roundedMinutes],
    0,
  );
  const projectName = data[0][ColIndex.projectName] || "Development";

  // Replace placeholders in the document body and header
  replaceText(body, Dates, { companyName, totalHours, projectName });
  if (header)
    replaceText(header, Dates, { companyName, totalHours, projectName });

  // Find the table placeholder in the document to insert data
  const tables = body.getTables();
  let placeholderTable = null;
  let placeholderRowIndex = -1;

  // Search for the placeholder '{{tableData}}' in the tables
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    const numRows = table.getNumRows();
    for (let j = 0; j < numRows; j++) {
      const row = table.getRow(j);
      if (row.getCell(0).getText().indexOf("{{tableData}}") !== -1) {
        placeholderTable = table;
        placeholderRowIndex = j;
        break;
      }
    }
    if (placeholderTable) break;
  }

  // Throw an error if the placeholder is not found
  if (placeholderRowIndex === -1) {
    throw new Error("Placeholder '{{tableData}}' not found in the document.");
  }

  // Loop through the data and insert rows into the table
  data.forEach((rowData, index) => {
    const formattedDate = Utilities.formatDate(
      new Date(rowData[ColIndex.date]),
      Session.getScriptTimeZone(),
      "MM/dd/yy",
    );
    const newRow = placeholderTable.insertTableRow(
      placeholderRowIndex + 1 + index,
    );

    // Create cells for date, project name, title, description, and minutes
    const dateCell = newRow.appendTableCell(formattedDate);
    //const projectNameCell = newRow.appendTableCell(rowData[ColIndex.projectName]);
    const titleCell = newRow.appendTableCell(rowData[ColIndex.title]);
    const descriptionCell = newRow.appendTableCell();
    const minutesCell = newRow.appendTableCell(
      rowData[ColIndex.roundedMinutes],
    );

    // Set alignment for date and minutes cells
    dateCell
      .getChild(0)
      .asParagraph()
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    minutesCell
      .getChild(0)
      .asParagraph()
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    // Insert description with link or as HTML content
    const link = rowData[ColIndex.links];
    const descriptionText = rowData[ColIndex.description];
    if (link && isHyperlink(link)) {
      // Insert description text with hyperlink
      let textElement = descriptionCell
        .appendParagraph("")
        .appendText(descriptionText || "");
      textElement.setLinkUrl(link); // Set the link for the whole text
    } else {
      // Insert description as formatted content (supports Markdown, HTML, or plain text)
      insertContentInDoc(descriptionCell, descriptionText || "");
    }
  });

  // Remove the placeholder row after inserting data
  placeholderTable.removeRow(placeholderRowIndex);
  docFile.saveAndClose();
};
