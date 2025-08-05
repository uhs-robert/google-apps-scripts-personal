// time-tracker/export-to-pdf.js

// === Config ================================================
const templateDocId = "1osZLyS7V_hUfdIX50xivWMlSmNAedx8w6ydt7b6W_R0"; // Replace with your template doc ID
const destFolderId = "1_bgB_uwwJ5DlYP6fG35sIFSwZC4Yxhjj"; // Replace with your folder ID

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
const Font = {
  size: {
    base: 9.5,
    code: 8.5,
  },
};

// === Styles ================================================
const Styles = {
  h1: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base + 1,
    [DocumentApp.Attribute.BOLD]: true,
  },
  h2: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base + 0.5,
    [DocumentApp.Attribute.BOLD]: true,
  },
  h3: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base + 0.3,
    [DocumentApp.Attribute.BOLD]: true,
  },
  h4: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base,
    [DocumentApp.Attribute.BOLD]: true,
  },
  h5: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.ITALIC]: true,
  },
  h6: {
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base,
    [DocumentApp.Attribute.ITALIC]: true,
  },
  body: {
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]:
      DocumentApp.HorizontalAlignment.LEFT,
    [DocumentApp.Attribute.FONT_FAMILY]: "Calibri",
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.base,
    [DocumentApp.Attribute.BOLD]: false,
  },
  codeInline: {
    [DocumentApp.Attribute.FONT_FAMILY]: "Courier New",
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.code,
    [DocumentApp.Attribute.BACKGROUND_COLOR]: "#efefef",
  },
  codeBlock: {
    [DocumentApp.Attribute.FONT_FAMILY]: "Courier New",
    [DocumentApp.Attribute.FONT_SIZE]: Font.size.code,
    [DocumentApp.Attribute.BACKGROUND_COLOR]: "#efefef",
    [DocumentApp.Attribute.INDENT_START]: 18,
    [DocumentApp.Attribute.INDENT_END]: 18,
    [DocumentApp.Attribute.SPACING_BEFORE]: 6,
    [DocumentApp.Attribute.SPACING_AFTER]: 6,
  },
};

// === Utilities ================================================
/**
 * Creates standardized filename for time tracking reports
 * @param {Object} Dates - Date information object from getDates()
 * @param {string} companyName - Company name for filename
 * @returns {string} Formatted filename without extension
 */
const generateFileName = (Dates, companyName) => {
  return `${companyName}_${Dates.formattedStart}_to_${Dates.formattedEnd}_Time_Tracking_Report`;
};

/**
 * Retrieves time entries matching company and project criteria
 * @param {string} companyName - Company name filter
 * @param {string} projectName - Project name filter
 * @returns {Array[]} Filtered rows from time tracking sheet
 */
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

/**
 * P,erforms template placeholder replacement in document sections
 * @param {GoogleAppsScript.Document.Body|GoogleAppsScript.Document.HeaderSection} source - Document section to update
 * @param {Object} Dates - Date formatting object
 * @param {Object} Info - Company and project information
 */
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

// Function to append headers with specific formatting
/**
 * Applies Google Docs heading styles to paragraphs
 * @param {GoogleAppsScript.Document.Paragraph} paragraph - Target paragraph
 * @param {string} text - Header text content
 * @param {string} headerType - HTML header tag (h1-h6)
 */
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
  paragraph.setSpacingBefore(6);
  paragraph.appendText(text).setAttributes(Styles[headerType]);
};

// Function to append styled text with specific formatting
/**
 * Applies text formatting to paragraph content
 * @param {GoogleAppsScript.Document.Paragraph} paragraph - Target paragraph
 * @param {string} text - Text content to style
 * @param {Object} styles - Formatting options (bold, italic, underline)
 */
const appendStyledText = (paragraph, text, styles) => {
  let textElement = paragraph.appendText(text).setAttributes(Styles.body);
  if (styles.bold) textElement.setBold(true);
  if (styles.italic) textElement.setItalic(true);
  if (styles.underline) textElement.setUnderline(true);
};

// Functions for processing lists and list items
/**
 * Converts HTML lists to Google Docs list formatting
 * @param {GoogleAppsScript.Document.Element} parent - Container element
 * @param {GoogleAppsScript.XML_Service.Element} listElement - Source list element
 * @param {boolean} isUnordered - List type (bullet vs numbered)
 * @param {number} nestLevel - Indentation level
 */
const processList = (parent, listElement, isUnordered, nestLevel = 0) => {
  listElement.getChildren().forEach((child) => {
    if (child.getName() === "li") {
      appendListItem(parent, child.getText(), isUnordered, nestLevel);
    } else if (child.getName() === "ul" || child.getName() === "ol") {
      processList(parent, child, isUnordered, nestLevel++);
    }
  });
};

/**
 * Creates formatted list items with proper indentation
 * @param {GoogleAppsScript.Document.Element} parent - Container element
 * @param {string} text - List item text content
 * @param {boolean} isUnordered - List type (bullet vs numbered)
 * @param {number} nestLevel - Indentation level
 */
const appendListItem = (parent, text, isUnordered, nestLevel = 0) => {
  const listItem = parent.appendListItem(text);
  listItem.setAttributes(Styles.body);
  listItem.setGlyphType(
    isUnordered ? DocumentApp.GlyphType.BULLET : DocumentApp.GlyphType.NUMBER,
  );
  // listItem.editAsText().setFontSize(Font.size.base);
  if (nestLevel.length > 0) {
    const textElement = listItem.editAsText();
    textElement.insertText(0, "\t".repeat(nestLevel));
  }
};

/**
 * Processes Markdown tokens and converts them to Google Docs formatting
 * @param {GoogleAppsScript.Document.TableCell} docCell - Target table cell
 * @param {Array} tokens - Parsed Markdown tokens
 */
const processMarkdownTokens = (docCell, tokens) => {
  docCell.clear();
  docCell.setAttributes(Styles.body);

  tokens.forEach((token) => {
    switch (token.type) {
      case "header":
        const headerParagraph = docCell.appendParagraph("");
        appendHeader(headerParagraph, token.content, `h${token.level}`);
        break;

      case "paragraph":
        if (token.content != "") {
          const paragraph = docCell.appendParagraph("");
          paragraph.setAttributes(Styles.body);
          appendInlineTokens(paragraph, token.content);
        }
        break;

      case "list":
        token.items.forEach((item) => {
          const listItem = docCell.appendListItem("");
          listItem.setGlyphType(
            token.ordered
              ? DocumentApp.GlyphType.NUMBER
              : DocumentApp.GlyphType.BULLET,
          );
          listItem.editAsText().setFontSize(Font.size.base);

          // Handle indentation
          if (item.indent > 0) {
            const textElement = listItem.editAsText();
            textElement
              .insertText(0, "\t".repeat(item.indent))
              .setAttributes(Styles.body);
          }

          // Add inline content to the list item text
          appendInlineTokensToText(listItem.editAsText(), item.content);
        });
        break;

      case "blockquote":
        const quoteParagraph = docCell.appendParagraph("");
        quoteParagraph.setIndentStart(36); // Indent blockquotes
        appendInlineTokens(quoteParagraph, token.content);
        quoteParagraph.setFontSize(Font.size.base);
        break;

      case "codeblock":
        const codeParagraph = docCell.appendParagraph(token.content);
        codeParagraph.setAttributes(Styles.codeBlock);
        break;
    }
  });

  // Remove empty first paragraph if it exists
  if (docCell.getNumChildren() > 0) {
    const firstChild = docCell.getChild(0);
    if (
      firstChild.asParagraph &&
      firstChild.asParagraph().getText().trim() === ""
    ) {
      docCell.removeChild(firstChild);
    }
  }
};

/**
 * Processes inline Markdown tokens and appends them to a paragraph
 * @param {GoogleAppsScript.Document.Paragraph} paragraph - Target paragraph
 * @param {Array} inlineTokens - Array of inline formatting tokens
 */
const appendInlineTokens = (paragraph, inlineTokens) => {
  inlineTokens.forEach((token) => {
    switch (token.type) {
      case "text":
        const textElement = paragraph.appendText(token.text);
        textElement.setAttributes(Styles.body);
        break;

      case "bold":
        const boldText = paragraph.appendText(token.text);
        boldText.setAttributes(Styles.body);
        boldText.setBold(true);
        break;

      case "italic":
        const italicText = paragraph.appendText(token.text);
        italicText.setAttributes(Styles.body);
        italicText.setItalic(true);
        break;

      case "strikethrough":
        const strikeText = paragraph.appendText(token.text);
        strikeText.setAttributes(Styles.body);
        strikeText.setStrikethrough(true);
        break;

      case "code":
        const codeText = paragraph.appendText(token.text);
        codeText.setAttributes(Styles.codeInline);
        break;

      case "link":
        const linkText = paragraph.appendText(token.text);
        linkText.setAttributes(Styles.body);
        linkText.setLinkUrl(token.url);
        break;
    }
  });
};

/**
 * Processes inline Markdown tokens and applies them to a Text element
 * @param {GoogleAppsScript.Document.Text} textElement - Target text element
 * @param {Array} inlineTokens - Array of inline formatting tokens
 */
const appendInlineTokensToText = (textElement, inlineTokens) => {
  let currentIndex = 0;
  const codeRanges = [];

  inlineTokens.forEach((token) => {
    const startIndex = currentIndex;
    const text = token.text || "";

    switch (token.type) {
      case "text":
        textElement.insertText(currentIndex, text);
        textElement.setAttributes(
          startIndex,
          currentIndex + text.length - 1,
          Styles.body,
        );
        currentIndex += text.length;
        break;

      case "bold":
        textElement.insertText(currentIndex, text);
        textElement.setAttributes(
          startIndex,
          currentIndex + text.length - 1,
          Styles.body,
        );
        textElement.setBold(startIndex, currentIndex + text.length - 1, true);
        currentIndex += text.length;
        break;

      case "italic":
        textElement.insertText(currentIndex, text);
        textElement.setAttributes(
          startIndex,
          currentIndex + text.length - 1,
          Styles.body,
        );
        textElement.setItalic(startIndex, currentIndex + text.length - 1, true);
        currentIndex += text.length;
        break;

      case "strikethrough":
        textElement.insertText(currentIndex, text);
        textElement.setAttributes(
          startIndex,
          currentIndex + text.length - 1,
          Styles.body,
        );
        textElement.setStrikethrough(
          startIndex,
          currentIndex + text.length - 1,
          true,
        );
        currentIndex += text.length;
        break;

      case "code":
        textElement.insertText(currentIndex, text);
        textElement.setFontFamily(
          startIndex,
          currentIndex + text.length - 1,
          "Courier New",
        );
        textElement.setFontSize(
          startIndex,
          currentIndex + text.length - 1,
          Font.size.code,
        );
        textElement.setBackgroundColor(
          startIndex,
          currentIndex + text.length - 1,
          "#efefef",
        );
        codeRanges.push({
          start: startIndex,
          end: currentIndex + text.length - 1,
        });
        currentIndex += text.length;
        break;

      case "link":
        textElement.insertText(currentIndex, text);
        textElement.setAttributes(
          startIndex,
          currentIndex + text.length - 1,
          Styles.body,
        );
        textElement.setLinkUrl(
          startIndex,
          currentIndex + text.length - 1,
          token.url,
        );
        currentIndex += text.length;
        break;
    }
  });

  // After all text is added, ensure only code ranges have background color
  const totalLength = textElement.getText().length;
  if (codeRanges.length > 0 && totalLength > 0) {
    // Clear background for entire text first
    textElement.setBackgroundColor(0, totalLength - 1, null);
    // Then re-apply background only to code ranges
    codeRanges.forEach((range) => {
      textElement.setBackgroundColor(range.start, range.end, "#efefef");
    });
  }
};

/**
 * Processes content (Markdown, HTML, or plain text) for insertion into document cells
 * @param {GoogleAppsScript.Document.TableCell} docCell - Target table cell
 * @param {string} content - Content to process and insert
 */
const insertContentInDoc = (docCell, content) => {
  docCell.setAttributes(Styles.body);
  if (isMarkdown(content)) {
    const tokens = parseMarkdown(content);
    processMarkdownTokens(docCell, tokens);
    // docCell.setFontSize(Font.size.body);
  } else if (isHtml(content)) {
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
    docCell.editAsText().setAttributes(Styles.body);
  }
};

/**
 * Removes unsafe HTML tags while preserving formatting elements
 * @param {string} html - Raw HTML content
 * @returns {string} Sanitized HTML with only allowed tags
 */
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

/**
 * Recursively converts XML/HTML elements to Google Docs formatting
 * @param {GoogleAppsScript.Document.Element} parent - Target document element
 * @param {GoogleAppsScript.XML_Service.Element} element - Source XML element
 * @param {Object} styles - Current text styling state
 */
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
