// time-tracker/ExportFunction.js
const ColIndex = {
  payRate: 0,
  companyName: 1,
  projectName: 2,
  title: 3,
  description: 4,
  links: 5,
  status: 6,
  startTime: 7,
  minutes: 8,
  roundedMinutes: 9,
  date: 10,
  week: 11,
  month: 12,
  year: 13,
};
const FontSize = 9.5;

const exportToPDF = (companyName, projectName, log = false) => {
  try {
    const templateDocId = "1osZLyS7V_hUfdIX50xivWMlSmNAedx8w6ydt7b6W_R0"; // Replace with your template doc ID
    const destFolderId = "1r6mK9x7ZXZfWYLgSNaXRf4whG5FB1sRf"; // Replace with your folder ID
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

const getOrCreateFolder = (parentFolder, folderName) => {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext()
    ? folders.next()
    : parentFolder.createFolder(folderName);
};

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

const generateFileName = (Dates, companyName) => {
  return `${companyName}_${Dates.formattedStart}_to_${Dates.formattedEnd}_Time_Tracking_Report`;
};

const createDocumentFromTemplate = (templateDocId, fileName, destFolder) => {
  const doc = DriveApp.getFileById(templateDocId).makeCopy(
    fileName,
    destFolder,
  );
  const docId = doc.getId();
  const docFile = DocumentApp.openById(docId);
  return { docFile, docId };
};

const convertDocToPDF = (docId, fileName, destFolder) => {
  const pdf = DriveApp.getFileById(docId).getAs("application/pdf");
  const pdfFile = destFolder.createFile(pdf).setName(`${fileName}.pdf`);
  return pdfFile;
};

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

const getFilteredData = (companyName, projectName) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("🕑 Time");
  const dataRange = sheet.getRange(`B3:O${sheet.getLastRow()}`);
  const data = dataRange.getValues();

  return data.filter((row) => {
    const rowCompanyName = row[1]; // Company Name is in column C (index 2)
    const rowProjectName = row[2]; // Project Name is in column D (index 3)

    return rowCompanyName === companyName && rowProjectName === projectName;
  });
};

const replaceText = (source, Dates, Info) => {
  const replacements = {
    "{{companyName}}": Info.companyName,
    "{{totalHours}}": Info.totalHours,
    "{{projectName}}": Info.projectName,
    "{{startDate}}": Dates.formattedStart,
    "{{endDate}}": Dates.formattedEnd,
    "{{todayLong}}": Dates.todayLong,
    "{{todayShort}}": Dates.todayShort,
  };

  Object.keys(replacements).forEach((key) => {
    source.replaceText(key, replacements[key]);
  });
};
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
      // Insert description as HTML content if no link is present
      insertHtmlContentInDoc(descriptionCell, descriptionText || "");
    }
  });

  // Remove the placeholder row after inserting data
  placeholderTable.removeRow(placeholderRowIndex);
  docFile.saveAndClose();
};

const insertHtmlContentInDoc = (docCell, content) => {
  if (isHtml(content)) {
    const sanitizedContent = sanitizeHtml(content);
    const parser = XmlService.parse("<div>" + sanitizedContent + "</div>");
    const root = parser.getRootElement();
    Logger.log({ sanitizedContent });
    docCell.clear();
    processElement(docCell, root);
    // Remove the first paragraph if it exists
    if (docCell.getNumChildren() > 0) {
      const firstChild = docCell.getChild(0);
      //const lastChild = docCell.getChild(docCell.getNumChildren() - 1);
      if (firstChild.asParagraph) docCell.removeChild(firstChild);
      //if(lastChild.asParagraph) docCell.removeChild(lastChild);
    }
  } else {
    docCell.setText(content);
  }
};

// Helper function to detect if a string contains a hyperlink
const isHyperlink = (text) => {
  const urlPattern = /(https?:\/\/[^\s]+)/g;
  return urlPattern.test(text);
};

// Extract link text and URL from the content
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

const isHtml = (content) => {
  return /<\/?[a-z][\s\S]*>/i.test(content);
};

const sanitizeHtml = (html) => {
  const allowedTags = [
    "b",
    "i",
    "u",
    "ul",
    "ol",
    "li",
    "div",
    "p",
    "br",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
  ];
  return html.replace(/<\/?([a-z][a-z0-9]*)\b[^>]*>/gi, (match, tag) => {
    return allowedTags.includes(tag.toLowerCase()) ? match : "";
  });
};

const processElement = (
  parent,
  element,
  styles = { bold: false, underline: false, italic: false },
) => {
  let currentParagraph = null;
  const children = element.getChildren();

  if (children.length <= 0) {
    const text = element.getText();
    if (parent.asParagraph) parent.appendText(text);
    return;
  }

  children.forEach((child) => {
    const text = child.getText() || "";
    const type = child.getName();

    switch (type) {
      case "b":
        styles.bold = true;
      case "i":
        if (!styles.bold) styles.italic = true;
      case "u":
        if (!styles.bold && !styles.italic) styles.underline = true;
        if (!currentParagraph) currentParagraph = parent.appendParagraph("");
        appendStyledText(currentParagraph, text, styles);
        styles = { bold: false, underline: false, italic: false };
        break;
      case "p":
      case "div":
        currentParagraph = parent.appendParagraph("");
        processElement(currentParagraph, child, styles);
        break;
      case "h1":
      case "h2":
      case "h3":
      case "h4":
      case "h5":
      case "h6":
        if (currentParagraph && currentParagraph.getText().trim())
          parent.appendParagraph("");
        currentParagraph = parent.appendParagraph("");
        appendHeader(currentParagraph, text, type);
        break;
      case "ul":
      case "ol":
        if (currentParagraph) currentParagraph = null;
        processList(parent, child, type === "ul");
        break;
      case "li":
        appendListItem(parent, text, parent.isUnordered);
        break;
      case "br":
        if (currentParagraph) currentParagraph.appendText("\n");
        break;
      default:
        if (text) {
          if (!currentParagraph) {
            currentParagraph = parent.appendParagraph(text);
          } else {
            currentParagraph.appendText(" " + text);
          }
        }
    }
    // Reset styles for inline elements, not for block-level containers
    if (!["p", "div"].includes(type))
      styles = { bold: false, underline: false, italic: false };
  });
};

// Function to append headers with specific formatting
const appendHeader = (paragraph, text, headerType) => {
  const headers = {
    h1: DocumentApp.ParagraphHeading.HEADING1,
    h2: DocumentApp.ParagraphHeading.HEADING2,
    h3: DocumentApp.ParagraphHeading.HEADING3,
    h4: DocumentApp.ParagraphHeading.HEADING4,
    h5: DocumentApp.ParagraphHeading.HEADING5,
    h6: DocumentApp.ParagraphHeading.HEADING6,
  };
  paragraph.setHeading(headers[headerType]);
  paragraph.appendText(text);
};

// Function to append styled text with specific formatting
const appendStyledText = (paragraph, text, styles) => {
  let textElement = paragraph.appendText(text);
  if (styles.bold) textElement.setBold(true);
  if (styles.italic) textElement.setItalic(true);
  if (styles.underline) textElement.setUnderline(true);
};

// Functions for processing lists and list items
const processList = (parent, listElement, isUnordered, nestLevel = 0) => {
  listElement.getChildren().forEach((child) => {
    if (child.getName() === "li") {
      appendListItem(parent, child.getText(), isUnordered, nestLevel);
    } else if (child.getName() === "ul" || child.getName() === "ol") {
      processList(parent, child, isUnordered, nestLevel++);
    }
  });
};

const appendListItem = (parent, text, isUnordered, nestLevel = 0) => {
  const listItem = parent.appendListItem(text);
  listItem.setGlyphType(
    isUnordered ? DocumentApp.GlyphType.BULLET : DocumentApp.GlyphType.NUMBER,
  );
  listItem.editAsText().setFontSize(FontSize);
  if (nestLevel.length > 0) {
    const textElement = listItem.editAsText();
    textElement.insertText(0, "\t".repeat(nestLevel));
  }
};
