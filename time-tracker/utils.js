// time-tracker/utils.js

// === Google Utilities ================================================
/**
 * Ensures company-specific folder exists in Drive destination
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - Parent Drive folder
 * @param {string} folderName - Name of folder to create or find
 * @returns {GoogleAppsScript.Drive.Folder} The company folder
 */
const getOrCreateFolder = (parentFolder, folderName) => {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext()
    ? folders.next()
    : parentFolder.createFolder(folderName);
};

/**
 * Duplicates template document for customization
 * @param {string} templateDocId - Google Doc template ID
 * @param {string} fileName - Name for the new document
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder
 * @returns {Object} Document file and ID for further processing
 */
const createDocumentFromTemplate = (templateDocId, fileName, destFolder) => {
  const doc = DriveApp.getFileById(templateDocId).makeCopy(
    fileName,
    destFolder,
  );
  const docId = doc.getId();
  const docFile = DocumentApp.openById(docId);
  return { docFile, docId };
};

/**
 * Converts Google Doc to PDF in specified folder
 * @param {string} docId - Document ID to convert
 * @param {string} fileName - Base filename for PDF
 * @param {GoogleAppsScript.Drive.Folder} destFolder - PDF destination folder
 * @returns {GoogleAppsScript.Drive.File} Created PDF file
 */
const convertDocToPDF = (docId, fileName, destFolder) => {
  const pdf = DriveApp.getFileById(docId).getAs("application/pdf");
  const pdfFile = destFolder.createFile(pdf).setName(`${fileName}.pdf`);
  return pdfFile;
};

/**
 * Displays styled HTML dialog in Sheets UI
 * @param {string} title - Dialog window title
 * @param {string} message - HTML content for dialog body
 */
const showDialog = (title, message) => {
  const htmlContent = `
    <html>
      <head>
        <style>
          body {
            font-family: Monaco, monospace;
          }
          p {
            margin: 10px 0;
            font-size: 14px;
          }
          a {
            color: #1a73e8;
            text-decoration: none;
          }
          a:hover {
            text-decoration: underline;
          }
        </style>
      </head>
      <body>
        <div>
          ${message}
        </div>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
};

// === General Utilities ================================================
/**
 * Extracts and formats date range information from time tracking data
 * @param {Array[]} data - Filtered time tracking rows
 * @returns {Object} Date object with formatted strings for template replacement
 */
const getDates = (data) => {
  const today = new Date();
  const actualStart = new Date(data[0][ColIndex.date]);
  const actualEnd = new Date(data[data.length - 1][ColIndex.date]);
  const timeZone = Session.getScriptTimeZone();

  return {
    today,
    todayShort: Utilities.formatDate(today, timeZone, "MM-dd-yyyy"),
    todayLong: Utilities.formatDate(today, timeZone, "MMMM dd, yyyy"),
    actualStart,
    actualEnd,
    formattedStart: Utilities.formatDate(actualStart, timeZone, "MM-dd-yyyy"),
    formattedEnd: Utilities.formatDate(actualEnd, timeZone, "MM-dd-yyyy"),
  };
};

// Helper function to detect if a string contains a hyperlink
/**
 * Detects if text contains HTTP/HTTPS URLs
 * @param {string} text - Text to analyze
 * @returns {boolean} True if hyperlink found
 */
const isHyperlink = (text) => {
  const urlPattern = /(https?:\/\/[^\s]+)/g;
  return urlPattern.test(text);
};

// Extract link text and URL from the content
/**
 * Extracts URL and link text from mixed content
 * @param {string} text - Text potentially containing URLs
 * @returns {Object} Separated link text and URL components
 */
const extractLinkData = (text) => {
  const urlPattern = /(https?:\/\/[^\s]+)/g;
  const urlMatch = text.match(urlPattern);
  if (urlMatch && urlMatch.length > 0) {
    const url = urlMatch[0];
    const linkText = text.replace(url, "").trim();
    return { text: linkText || url, url };
  }
  return { text, url: "" };
};

/**
 * Identifies HTML content by detecting tag patterns
 * @param {string} content - Content to analyze
 * @returns {boolean} True if HTML tags detected
 */
const isHtml = (content) => {
  return /<\/?[a-z][\s\S]*>/i.test(content);
};
