/**
 * Functions for retrieving and processing invoice data from the Invoices sheet
 */

/**
 * Gets all invoices from the Invoices sheet
 * @returns {Array<Array>} Array of invoice rows (excluding header and empty rows)
 */
function getAllInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVOICES_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${INVOICES_SHEET_NAME}" not found`);
  }

  const data = sheet.getDataRange().getValues();

  // Return all rows except header (row 0) and filter out empty rows
  // A row is considered empty if the Client column (index 1) is empty
  return data.slice(1).filter(row => row[InvoiceColIndex.CLIENT] && String(row[InvoiceColIndex.CLIENT]).trim() !== '');
}

/**
 * Gets a single invoice by row number
 * @param {number} rowNumber - The row number in the sheet (1-indexed, excluding header)
 * @returns {Object} Invoice data object
 */
function getInvoiceByRow(rowNumber) {
  const invoices = getAllInvoices();

  // rowNumber is 1-indexed (first data row = 1), but array is 0-indexed
  const rowIndex = rowNumber - 1;

  if (rowIndex < 0 || rowIndex >= invoices.length) {
    throw new Error(`Invalid row number: ${rowNumber}`);
  }

  const row = invoices[rowIndex];

  // Log raw values for debugging
  Logger.log('Raw row data - HOURS_TOTAL (col ' + InvoiceColIndex.HOURS_TOTAL + '): ' + row[InvoiceColIndex.HOURS_TOTAL]);
  Logger.log('Raw row data - RATE (col ' + InvoiceColIndex.RATE + '): ' + row[InvoiceColIndex.RATE]);
  Logger.log('Raw row data - GROSS (col ' + InvoiceColIndex.GROSS + '): ' + row[InvoiceColIndex.GROSS]);

  return {
    rowNumber: rowNumber,
    year: row[InvoiceColIndex.YEAR],
    client: row[InvoiceColIndex.CLIENT],
    projectTitle: row[InvoiceColIndex.PROJECT_TITLE],
    active: row[InvoiceColIndex.ACTIVE],
    source: row[InvoiceColIndex.SOURCE],
    monthName: row[InvoiceColIndex.MONTH_NAME],
    start: row[InvoiceColIndex.START],
    end: row[InvoiceColIndex.END],
    recur: row[InvoiceColIndex.RECUR],
    status: row[InvoiceColIndex.STATUS],
    hoursTotal: row[InvoiceColIndex.HOURS_TOTAL],
    hoursAdj: row[InvoiceColIndex.HOURS_ADJ],
    hoursUsed: row[InvoiceColIndex.HOURS_USED],
    hoursRem: row[InvoiceColIndex.HOURS_REM],
    percDone: row[InvoiceColIndex.PERC_DONE],
    progress: row[InvoiceColIndex.PROGRESS],
    pif: row[InvoiceColIndex.PIF],
    rate: row[InvoiceColIndex.RATE],
    gross: row[InvoiceColIndex.GROSS],
    tax: row[InvoiceColIndex.TAX],
    net: row[InvoiceColIndex.NET],
    report: row[InvoiceColIndex.REPORT],
    notes: row[InvoiceColIndex.NOTES]
  };
}

/**
 * Generates a custom invoice number in format: YYYY-XXYY-RRRR or YYYY-XXYY-RRRR-PF
 * Where:
 *   YYYY = Year
 *   XX = First letter of client name as number (A=01, B=02, ..., Z=26)
 *   YY = Second letter of client name as number
 *   RRRR = Row number padded to 4 digits
 *   -PF = Processing fee suffix (only if includeFee is true)
 *
 * @param {number} year - The invoice year
 * @param {string} clientName - The client name
 * @param {number} rowNumber - The row number
 * @param {boolean} includeFee - Whether to include processing fee suffix
 * @returns {string} Generated invoice number
 */
function generateInvoiceNumber(year, clientName, rowNumber, includeFee = false) {
  // Convert first two letters of client name to numbers
  const getLetterNumber = (char) => {
    const upperChar = char.toUpperCase();
    const code = upperChar.charCodeAt(0);

    // A=65 in ASCII, so A=1, B=2, etc.
    if (code >= 65 && code <= 90) {
      return String(code - 64).padStart(2, '0');
    }

    // If not a letter, return 00
    return '00';
  };

  const firstLetter = clientName.length > 0 ? getLetterNumber(clientName[0]) : '00';
  const secondLetter = clientName.length > 1 ? getLetterNumber(clientName[1]) : '00';
  const paddedRow = String(rowNumber).padStart(4, '0');

  let invoiceNumber = `${year}-${firstLetter}${secondLetter}-${paddedRow}`;

  // Add processing fee suffix if needed
  if (includeFee) {
    invoiceNumber += INVOICE_NUMBER_FEE_SUFFIX;
  }

  return invoiceNumber;
}

/**
 * Gets a list of invoices formatted for display in the dialog
 * Returns newest invoices first (reversed order)
 * @returns {Array<Object>} Array of objects with display info
 */
function getInvoicesForDialog() {
  const invoices = getAllInvoices();

  // Map invoices to display objects with proper row numbers
  const invoiceList = invoices.map((row, index) => {
    const rowNumber = index + 1;
    const client = row[InvoiceColIndex.CLIENT];
    const projectTitle = row[InvoiceColIndex.PROJECT_TITLE];
    const year = row[InvoiceColIndex.YEAR];
    const monthName = row[InvoiceColIndex.MONTH_NAME];
    const hoursUsed = row[InvoiceColIndex.HOURS_USED];
    const gross = row[InvoiceColIndex.GROSS];

    const invoiceNumber = generateInvoiceNumber(year, client, rowNumber);

    return {
      rowNumber: rowNumber,
      invoiceNumber: invoiceNumber,
      displayText: `${invoiceNumber} - ${client} - ${projectTitle} (${monthName} ${year}) - $${gross}`,
      client: client,
      project: projectTitle,
      amount: gross
    };
  });

  // Reverse the array so newest invoices appear first, filtering out PIF invoices
  return invoiceList.filter(inv => !invoices[inv.rowNumber - 1][InvoiceColIndex.PIF]).reverse();
}
