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

/**
 * Identifies Markdown content by detecting common Markdown patterns
 * @param {string} content - Content to analyze
 * @returns {boolean} True if Markdown patterns detected
 */
const isMarkdown = (content) => {
  const markdownPatterns = [
    /^#{1,6}\s+/m,          // Headers: # ## ### etc
    /\*\*.*?\*\*/,          // Bold: **text**
    /\*.*?\*/,              // Italic: *text*
    /~~.*?~~/,              // Strikethrough: ~~text~~
    /`.*?`/,                // Inline code: `code`
    /^\s*[-*+]\s+/m,        // Unordered lists: - * +
    /^\s*\d+\.\s+/m,        // Ordered lists: 1. 2.
    /\[.*?\]\(.*?\)/,       // Links: [text](url)
    /^>\s+/m,               // Blockquotes: > text
    /```[\s\S]*?```/,       // Code blocks: ```code```
  ];
  
  return markdownPatterns.some(pattern => pattern.test(content));
};

/**
 * Parses Markdown content into structured tokens for Google Docs conversion
 * @param {string} content - Markdown content to parse
 * @returns {Array} Array of tokens representing the parsed content
 */
const parseMarkdown = (content) => {
  const tokens = [];
  const lines = content.split('\n');
  let currentList = null;
  let inCodeBlock = false;
  let codeBlockContent = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Handle code blocks
    if (line.trim().startsWith('```')) {
      if (inCodeBlock) {
        // End code block
        tokens.push({
          type: 'codeblock',
          content: codeBlockContent.join('\n')
        });
        codeBlockContent = [];
        inCodeBlock = false;
      } else {
        // Start code block
        inCodeBlock = true;
      }
      continue;
    }
    
    if (inCodeBlock) {
      codeBlockContent.push(line);
      continue;
    }
    
    // Empty line - end current list and add paragraph break
    if (line.trim() === '') {
      if (currentList) {
        tokens.push(currentList);
        currentList = null;
      }
      tokens.push({ type: 'paragraph', content: '' });
      continue;
    }
    
    // Headers
    const headerMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headerMatch) {
      if (currentList) {
        tokens.push(currentList);
        currentList = null;
      }
      tokens.push({
        type: 'header',
        level: headerMatch[1].length,
        content: headerMatch[2]
      });
      continue;
    }
    
    // Lists
    const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.+)$/);
    if (listMatch) {
      const indent = listMatch[1].length;
      const isOrdered = /\d+\./.test(listMatch[2]);
      const content = listMatch[3];
      
      if (!currentList || currentList.ordered !== isOrdered) {
        if (currentList) tokens.push(currentList);
        currentList = {
          type: 'list',
          ordered: isOrdered,
          items: []
        };
      }
      
      currentList.items.push({
        content: parseInlineMarkdown(content),
        indent: Math.floor(indent / 2)
      });
      continue;
    }
    
    // Blockquotes
    const quoteMatch = line.match(/^>\s*(.+)$/);
    if (quoteMatch) {
      if (currentList) {
        tokens.push(currentList);
        currentList = null;
      }
      tokens.push({
        type: 'blockquote',
        content: parseInlineMarkdown(quoteMatch[1])
      });
      continue;
    }
    
    // Regular paragraph
    if (currentList) {
      tokens.push(currentList);
      currentList = null;
    }
    
    tokens.push({
      type: 'paragraph',
      content: parseInlineMarkdown(line)
    });
  }
  
  // Close any remaining list
  if (currentList) {
    tokens.push(currentList);
  }
  
  return tokens;
};

/**
 * Parses inline Markdown formatting within text
 * @param {string} text - Text with inline Markdown
 * @returns {Array} Array of inline tokens with formatting
 */
const parseInlineMarkdown = (text) => {
  const tokens = [];
  let remaining = text;
  
  // Define inline patterns with priority order
  const patterns = [
    { regex: /\[([^\]]+)\]\(([^)]+)\)/, type: 'link', groups: ['text', 'url'] },
    { regex: /`([^`]+)`/, type: 'code', groups: ['text'] },
    { regex: /\*\*([^*]+)\*\*/, type: 'bold', groups: ['text'] },
    { regex: /\*([^*]+)\*/, type: 'italic', groups: ['text'] },
    { regex: /~~([^~]+)~~/, type: 'strikethrough', groups: ['text'] },
  ];
  
  while (remaining.length > 0) {
    let matched = false;
    
    for (const pattern of patterns) {
      const match = remaining.match(pattern.regex);
      if (match && match.index === 0) {
        // Add the formatted content
        const token = { type: pattern.type };
        pattern.groups.forEach((group, index) => {
          token[group] = match[index + 1];
        });
        tokens.push(token);
        
        remaining = remaining.slice(match[0].length);
        matched = true;
        break;
      }
    }
    
    if (!matched) {
      // Find next special character or add rest as plain text
      let nextSpecialIndex = remaining.length;
      for (const pattern of patterns) {
        const match = remaining.match(pattern.regex);
        if (match && match.index < nextSpecialIndex) {
          nextSpecialIndex = match.index;
        }
      }
      
      if (nextSpecialIndex > 0) {
        tokens.push({
          type: 'text',
          text: remaining.slice(0, nextSpecialIndex)
        });
        remaining = remaining.slice(nextSpecialIndex);
      } else {
        tokens.push({
          type: 'text',
          text: remaining
        });
        remaining = '';
      }
    }
  }
  
  return tokens;
};

/**
 * Test function to validate Markdown parsing functionality
 * @returns {Object} Test results with sample conversions
 */
const testMarkdownParsing = () => {
  const testCases = [
    {
      name: 'Headers',
      input: '# Main Title\n## Subtitle\n### Section',
      expected: 'Should create heading1, heading2, heading3'
    },
    {
      name: 'Text Formatting',
      input: 'This is **bold** and *italic* and `code` text.',
      expected: 'Should apply formatting inline'
    },
    {
      name: 'Lists',
      input: '- Item 1\n- Item 2\n  - Nested item\n1. Numbered\n2. List',
      expected: 'Should create bullet and numbered lists'
    },
    {
      name: 'Links',
      input: 'Check out [Google](https://google.com) for search.',
      expected: 'Should create hyperlink'
    },
    {
      name: 'Code Block',
      input: '```\nfunction test() {\n  return "hello";\n}\n```',
      expected: 'Should create formatted code block'
    }
  ];
  
  const results = testCases.map(test => {
    const isDetected = isMarkdown(test.input);
    const tokens = isDetected ? parseMarkdown(test.input) : null;
    
    return {
      name: test.name,
      input: test.input,
      detected: isDetected,
      tokenCount: tokens ? tokens.length : 0,
      tokens: tokens,
      expected: test.expected
    };
  });
  
  Logger.log('Markdown Test Results:', results);
  return results;
};
