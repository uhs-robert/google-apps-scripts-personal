// time-tracker/document-generator.js

// === Document Generation ================================================
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
    const text = element.getText() || "";
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
