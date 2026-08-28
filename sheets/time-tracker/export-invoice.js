// sheets/time-tracker/export-invoice.js
/**
 * Main invoice generation and export functionality
 */

/**
 * Main function to export an invoice as PDF
 * @param {number} rowNumber - The row number of the invoice to export
 * @param {boolean} includeFee - Whether to include 3.5% processing fee
 * @returns {Object} Object containing success status and file URLs
 */
function exportInvoice(rowNumber, includeFee = false) {
  try {
    // Get invoice data
    const invoiceData = getInvoiceByRow(rowNumber);

    // Generate invoice number (with -PF suffix if fee included)
    const invoiceNumber = generateInvoiceNumber(
      invoiceData.year,
      invoiceData.client,
      invoiceData.rowNumber,
      includeFee,
    );

    // Get or create client folder
    const clientFolder = getOrCreateClientFolder(invoiceData.client);

    // Get the Invoices subfolder
    const invoicesFolder = getInvoicesSubfolder(clientFolder);

    // Generate filename (without extension)
    const filename = `${invoiceNumber} - ${invoiceData.client}`;

    // Create document from template
    const doc = createInvoiceDocument(invoiceNumber, invoiceData);
    const docId = doc.getId();

    // Update all placeholders
    updateInvoicePlaceholders(doc, invoiceData, invoiceNumber, includeFee);

    // CRITICAL: Must save and close to persist changes
    Logger.log("Saving and closing document...");
    doc.saveAndClose();
    Logger.log("Document saved and closed");

    // Wait for changes to fully propagate to Google Drive
    Utilities.sleep(3000);
    Logger.log("Converting to PDF...");

    // Convert to PDF
    const pdfFile = convertDocToPDF(docId, filename, invoicesFolder);
    Logger.log("PDF created successfully");

    // Delete the temporary document
    DriveApp.getFileById(docId).setTrashed(true);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      pdfFileId: pdfFile.getId(),
      invoiceNumber: invoiceNumber,
      client: invoiceData.client,
      invoiceData: invoiceData,
    };
  } catch (error) {
    Logger.log("Error exporting invoice: " + error.toString());
    throw error;
  }
}

/**
 * Gets or creates a client folder in the Clients parent folder
 * @param {string} clientName - The name of the client
 * @returns {Folder} The client folder
 */
function getOrCreateClientFolder(clientName) {
  const parentFolder = DriveApp.getFolderById(CLIENTS_PARENT_FOLDER_ID);

  // Check if client folder already exists
  const existingFolders = parentFolder.getFoldersByName(clientName);

  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  // Client folder doesn't exist, create it by copying the template
  const templateFolder = DriveApp.getFolderById(CLIENT_FOLDER_TEMPLATE_ID);

  // Copy the template folder
  const newFolder = copyFolder(templateFolder, clientName, parentFolder);

  return newFolder;
}

/**
 * Recursively copies a folder (Folder has no makeCopy method, unlike File)
 * @param {Folder} sourceFolder - The folder to copy
 * @param {string} newName - The name for the copied folder
 * @param {Folder} destinationParent - The parent folder for the copy
 * @returns {Folder} The newly created folder copy
 */
function copyFolder(sourceFolder, newName, destinationParent) {
  const newFolder = destinationParent.createFolder(newName);

  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), newFolder);
  }

  const folders = sourceFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    copyFolder(folder, folder.getName(), newFolder);
  }

  return newFolder;
}

/**
 * Gets the Invoices subfolder within a client folder
 * @param {Folder} clientFolder - The client folder
 * @returns {Folder} The Invoices subfolder
 */
function getInvoicesSubfolder(clientFolder) {
  const invoicesFolders = clientFolder.getFoldersByName("Invoices");

  if (invoicesFolders.hasNext()) {
    return invoicesFolders.next();
  }

  // If Invoices subfolder doesn't exist, create it
  return clientFolder.createFolder("Invoices");
}

/**
 * Creates a new document from the invoice template
 * @param {string} invoiceNumber - The invoice number
 * @param {Object} invoiceData - The invoice data
 * @returns {Document} The created document
 */
function createInvoiceDocument(invoiceNumber, invoiceData) {
  const templateFile = DriveApp.getFileById(INVOICE_TEMPLATE_ID);
  const tempName = `Invoice ${invoiceNumber} - ${invoiceData.client} (temp)`;

  // Make a copy of the template
  const docFile = templateFile.makeCopy(tempName);

  // Open the document for editing
  return DocumentApp.openById(docFile.getId());
}

/**
 * Updates all placeholders in the invoice document
 * @param {Document} doc - The document to update
 * @param {Object} invoiceData - The invoice data
 * @param {string} invoiceNumber - The generated invoice number
 * @param {boolean} includeFee - Whether to include processing fee
 */
function updateInvoicePlaceholders(
  doc,
  invoiceData,
  invoiceNumber,
  includeFee,
) {
  const body = doc.getBody();

  // Calculate dates
  const invoiceDate = new Date();
  const dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + DUE_DATE_DAYS);

  const timeZone = Session.getScriptTimeZone();
  const invoiceDateStr = Utilities.formatDate(
    invoiceDate,
    timeZone,
    "MMMM dd, yyyy",
  );
  const dueDateStr = Utilities.formatDate(dueDate, timeZone, "MMMM dd, yyyy");

  // Format currency values
  const formatCurrency = (value) => {
    if (typeof value === "number") {
      return value.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }
    return String(value);
  };

  // Calculate processing fee if needed
  // Customer pays the Gross amount (Net is your after-tax take home)
  let processingFee = 0;
  let totalDue = invoiceData.gross;

  if (includeFee) {
    processingFee = invoiceData.gross * PROCESSING_FEE_PERCENTAGE;
    totalDue = invoiceData.gross + processingFee;
  }

  // Build subject line
  const subject = `${invoiceData.projectTitle}: ${invoiceData.hoursTotal} hours`;

  // Create replacements object
  const replacements = {
    "{{invoice_number}}": invoiceNumber,
    "{{balance_due}}": formatCurrency(totalDue),
    "{{invoice_date}}": invoiceDateStr,
    "{{company_name}}": invoiceData.client,
    "{{terms}}": PAYMENT_TERMS,
    "{{due_date}}": dueDateStr,
    "{{subject}}": subject,
    "{{subtotal}}": formatCurrency(invoiceData.gross),
    "{{processing_fee}}": formatCurrency(processingFee),
    "{{total_due}}": formatCurrency(totalDue),
  };

  // Replace all text placeholders
  replaceInvoiceText(doc, replacements);

  // Handle processing fee row - delete it if no fee
  if (!includeFee) {
    deleteProcessingFeeRow(doc);
  }

  // Handle card payment links
  handleCardLinksReplacement(doc, includeFee);

  // Update the line item table
  updateLineItemTable(doc, invoiceData);
}

/**
 * Replaces text placeholders in the invoice document
 * @param {Document} doc - The document to update
 * @param {Object} replacements - Key-value pairs for text replacement
 */
function replaceInvoiceText(doc, replacements) {
  const body = doc.getBody();

  Logger.log(
    "Starting text replacement with " +
      Object.keys(replacements).length +
      " replacements",
  );

  // Replace in body
  Object.keys(replacements).forEach((key) => {
    // Escape special regex characters in the key (especially curly braces)
    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const result = body.replaceText(escapedKey, replacements[key]);
    Logger.log('Body replaced "' + key + '" -> "' + replacements[key] + '"');
  });

  // Replace in header (if it exists)
  try {
    const header = doc.getHeader();
    if (header) {
      Logger.log("Header section found, replacing text...");
      const headerText = header.getText();
      Logger.log("Full header text before replacement: " + headerText);

      // Get header structure
      const headerTables = header.getTables();
      const headerParagraphs = header.getParagraphs();
      Logger.log(
        "Header has " +
          headerTables.length +
          " tables and " +
          headerParagraphs.length +
          " paragraphs",
      );

      // Log all child elements in header
      const numChildren = header.getNumChildren();
      Logger.log("Header has " + numChildren + " child elements");
      for (let i = 0; i < numChildren; i++) {
        const child = header.getChild(i);
        const childType = child.getType();
        Logger.log("Header child[" + i + "] type: " + childType);

        // Try to get text from this child
        try {
          const childText = child.getText ? child.getText() : "N/A";
          Logger.log("Header child[" + i + "] text: " + childText);
        } catch (e) {
          Logger.log(
            "Header child[" + i + "] could not get text: " + e.message,
          );
        }
      }

      // Replace in header tables
      for (let t = 0; t < headerTables.length; t++) {
        const table = headerTables[t];
        const numRows = table.getNumRows();
        Logger.log("Header table " + t + " has " + numRows + " rows");

        for (let r = 0; r < numRows; r++) {
          const row = table.getRow(r);
          const numCells = row.getNumCells();

          for (let c = 0; c < numCells; c++) {
            const cell = row.getCell(c);
            const cellText = cell.getText();

            if (cellText) {
              Object.keys(replacements).forEach((key) => {
                if (cellText.includes(key)) {
                  Logger.log(
                    'Found "' +
                      key +
                      '" in header table[' +
                      t +
                      "] row[" +
                      r +
                      "] cell[" +
                      c +
                      "]",
                  );
                  const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                  cell.replaceText(escapedKey, replacements[key]);
                  Logger.log("Replaced with: " + replacements[key]);
                }
              });
            }
          }
        }
      }

      // Replace in header paragraphs (for text outside tables)
      for (let p = 0; p < headerParagraphs.length; p++) {
        const para = headerParagraphs[p];
        const paraText = para.getText();

        if (paraText) {
          Object.keys(replacements).forEach((key) => {
            if (paraText.includes(key)) {
              Logger.log('Found "' + key + '" in header paragraph[' + p + "]");
              const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
              para.replaceText(escapedKey, replacements[key]);
              Logger.log("Replaced with: " + replacements[key]);
            }
          });
        }
      }

      Logger.log("Header text after replacement: " + header.getText());
    } else {
      Logger.log("Header section is null");
    }
  } catch (e) {
    Logger.log("Error accessing header: " + e.message);
  }

  // Replace in footer (if it exists)
  try {
    const footer = doc.getFooter();
    if (footer) {
      Logger.log("Footer section found, replacing text...");
      Object.keys(replacements).forEach((key) => {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        footer.replaceText(escapedKey, replacements[key]);
        Logger.log(
          'Footer replaced "' + key + '" -> "' + replacements[key] + '"',
        );
      });
    } else {
      Logger.log("Footer section is null");
    }
  } catch (e) {
    // No footer section in document, skip
    Logger.log("Error accessing footer: " + e.message);
  }
}

/**
 * Updates the line item table in the invoice
 * @param {Document} doc - The document to update
 * @param {Object} invoiceData - The invoice data
 */
function updateLineItemTable(doc, invoiceData) {
  const body = doc.getBody();
  const tables = body.getTables();

  // Find the table row containing {{table_data}} placeholder
  let lineItemTable = null;
  let rowIndex = -1;

  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    const numRows = table.getNumRows();

    for (let r = 0; r < numRows; r++) {
      const row = table.getRow(r);
      const rowText = row.getText();

      // Look for {{table_data}} placeholder
      if (rowText.includes("{{table_data}}")) {
        lineItemTable = table;
        rowIndex = r;
        break;
      }
    }

    if (lineItemTable) break;
  }

  if (!lineItemTable || rowIndex === -1) {
    Logger.log(
      "Warning: Could not find {{table_data}} placeholder in any table",
    );
    return;
  }

  // Get the row to update
  const row = lineItemTable.getRow(rowIndex);

  // Log table structure
  const numCells = row.getNumCells();
  Logger.log("Table row has " + numCells + " cells");

  // Log current content of each cell
  for (let i = 0; i < numCells; i++) {
    const cellText = row.getCell(i).getText();
    Logger.log("Cell " + i + ' current text: "' + cellText + '"');
  }

  // Format currency
  const formatCurrency = (value) => {
    // Convert to number if it's a string
    const numValue = typeof value === "string" ? parseFloat(value) : value;

    if (typeof numValue === "number" && !isNaN(numValue)) {
      return (
        "$" +
        numValue.toLocaleString("en-US", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })
      );
    }
    return "$0.00";
  };

  // Log data for debugging
  Logger.log(
    "Invoice Data: hoursTotal=" +
      invoiceData.hoursTotal +
      ", rate=" +
      invoiceData.rate +
      ", gross=" +
      invoiceData.gross,
  );
  Logger.log("Formatted rate: " + formatCurrency(invoiceData.rate));
  Logger.log("Formatted gross: " + formatCurrency(invoiceData.gross));

  // Update cells - trying different cell mappings to find the right structure
  const qty = invoiceData.hoursTotal || invoiceData.hoursUsed || 0;
  const rateFormatted = formatCurrency(invoiceData.rate);
  const grossFormatted = formatCurrency(invoiceData.gross);

  // Try filling all 7 cells to see which ones appear where
  // Cell 0
  row.getCell(0).clear().setText("1");
  Logger.log("Cell 0 set to: 1");

  // Cell 1 - Item & Description with styled description
  const itemCell = row.getCell(1);
  itemCell.clear();
  const textElement = itemCell.editAsText();
  const fullText = LINE_ITEM_TITLE + "\n" + LINE_ITEM_DESCRIPTION;
  textElement.setText(fullText);
  textElement.setFontSize(0, LINE_ITEM_TITLE.length - 1, 11);
  const descStart = LINE_ITEM_TITLE.length + 1;
  textElement.setFontSize(descStart, fullText.length - 1, 9);
  textElement.setForegroundColor(descStart, fullText.length - 1, "#777777");
  Logger.log("Cell 1 set to: " + LINE_ITEM_TITLE);

  // Cell 2 and 3 are merged cells, proceed to 4
  row.getCell(4).clear().setText(String(qty));

  // Cell 5
  row.getCell(5).clear().setText(rateFormatted);
  Logger.log("Cell 4 set to: " + rateFormatted);
  // Cell 6
  row.getCell(6).clear().setText(grossFormatted);
  Logger.log("Cell 6 set to: " + grossFormatted);
}

/**
 * Deletes the table row containing the {{processing_fee}} placeholder
 * @param {Document} doc - The document to modify
 */
function deleteProcessingFeeRow(doc) {
  const body = doc.getBody();
  const tables = body.getTables();

  // Find and delete the row containing {{processing_fee}}
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    const numRows = table.getNumRows();

    for (let r = numRows - 1; r >= 0; r--) {
      const row = table.getRow(r);
      const rowText = row.getText();

      if (
        rowText.includes("{{processing_fee}}") ||
        rowText.includes("Processing Fee")
      ) {
        table.removeRow(r);
        Logger.log("Removed processing fee row at index " + r);
        return;
      }
    }
  }
}

/**
 * Replaces the {{card_links}} placeholder with payment links or disclaimer
 * @param {Document} doc - The document to modify
 * @param {boolean} includeFee - Whether to include payment links
 */
function handleCardLinksReplacement(doc, includeFee) {
  const body = doc.getBody();

  if (includeFee) {
    // Build payment links text
    const paymentLinksText = `Pay online via:\nStripe: ${STRIPE_PAYMENT_URL}\nPayPal: ${PAYPAL_PAYMENT_URL}`;

    // Replace placeholder with links
    body.replaceText("{{card_links}}", paymentLinksText);
  } else {
    // Replace with disclaimer
    body.replaceText("{{card_links}}", CARD_PAYMENT_DISCLAIMER);
  }
}

/**
 * Shows the invoice export dialog
 */
function showInvoiceExportDialog() {
  const html = HtmlService.createHtmlOutputFromFile("InvoiceDialog")
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, "Export Invoice");
}

/**
 * Handles the invoice export from the dialog
 * Called from the HTML dialog
 * @param {number} rowNumber - The row number to export
 * @param {boolean} includeFee - Whether to include processing fee
 * @returns {Object} Result object
 */
function handleInvoiceExport(rowNumber, includeFee = false) {
  try {
    const result = exportInvoice(rowNumber, includeFee);

    return {
      success: true,
      message: `Invoice ${result.invoiceNumber} generated successfully!`,
      pdfUrl: result.pdfUrl,
      pdfFileId: result.pdfFileId,
      invoiceNumber: result.invoiceNumber,
      rowNumber: rowNumber,
      includeFee: includeFee,
    };
  } catch (error) {
    return {
      success: false,
      message: "Error generating invoice: " + error.toString(),
    };
  }
}

/**
 * Sends the invoice email for an already-exported PDF.
 * Called separately after the user previews the PDF.
 * @param {number} rowNumber - Invoice row number
 * @param {boolean} includeFee - Whether processing fee was included
 * @param {string} pdfFileId - Drive file ID of the exported PDF
 * @param {string} pdfUrl - Drive URL for the View Invoice button
 * @returns {Object} Result object
 */
function handleSendInvoiceEmail(rowNumber, includeFee, pdfFileId, pdfUrl) {
  try {
    const invoiceData = getInvoiceByRow(rowNumber);
    const invoiceNumber = generateInvoiceNumber(
      invoiceData.year,
      invoiceData.client,
      rowNumber,
      includeFee,
    );
    const pdfBlob = DriveApp.getFileById(pdfFileId).getBlob();

    const draftUrl = sendInvoiceEmail(
      invoiceData,
      invoiceNumber,
      includeFee,
      pdfBlob,
      pdfUrl,
    );

    return { success: true, draftUrl: draftUrl };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
