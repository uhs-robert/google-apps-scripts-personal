// time-tracker/index.js

const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Export")
    .addItem("Export Time Entries to PDF", "showExportPDFDialog")
    .addToUi();
  ui.createMenu("Tools")
    .addItem("Wrap cell in HTML list", "wrapCellTextInHtmlList")
    .addToUi();
};

const onEdit = (e) => {
  const activeSheet = e.range.getSheet();
  const statusRange = e.source.getRangeByName("assignmentStatus");
  if (
    activeSheet.getName() !== statusRange.getSheet().getName() &&
    e.range.getColumn() !== statusRange.getColumn()
  )
    return;
  calculateTime({
    newStatus: e.value,
    prevStatus: e.oldValue,
    cellStartTime: e.range.offset(0, 1),
    cellElapsedTime: e.range.offset(0, 2),
  });
};

/* Manually wrap a cell's text in HTML list format
function wrapCellTextInHtmlList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();

  range.getValues().forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        // Split cell text by line breaks
        const lines = cellValue.split(/\r?\n/);

        // Only format if there is more than one line
        if (lines.length > 1) { 
          // Wrap lines in <ul><li></li></ul>
          const htmlList = `<ul>\n` + lines.map(line => `  <li>${line}</li>`).join('\n') + `\n</ul>`;
          range.getCell(rowIndex + 1, colIndex + 1).setValue(htmlList);
        }
      }
    });
  });
}
*/

function wrapCellTextInHtmlList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange(); // Get selected range

  range.getValues().forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (typeof cellValue === "string" && cellValue.trim() !== "") {
        const lines = cellValue.split(/\r?\n/); // Split lines by line breaks
        let htmlOutput = ""; // Final HTML output
        let stack = []; // Track nesting levels
        let prevLevel = 0; // Track the previous indentation level
        let isFirstLine = true; // Track if it's the first line

        lines.forEach((line) => {
          const trimmedLine = line.trim();

          // Detect the nesting level based on "-" prefixes
          const match = trimmedLine.match(/^(-+)\s*/); // Matches '-', '--', '---', etc.
          const level = match ? match[1].length : 0; // Count '-' as the nesting level

          if (isFirstLine && level === 0) {
            // Treat the first line as plain text if it has no '-' prefix
            htmlOutput += `<p>${trimmedLine}<p>\n`; // Plain text
            isFirstLine = false; // Mark first line processed
          } else {
            // Handle nested lists
            if (level > prevLevel) {
              // Open new nested lists
              for (let i = prevLevel; i < level; i++) {
                htmlOutput += `<ul>`;
                stack.push("ul");
              }
            } else if (level < prevLevel) {
              // Close nested lists
              for (let i = prevLevel; i > level; i--) {
                htmlOutput += `</ul>`;
                stack.pop();
              }
            }

            // Add list item, removing prefixes
            htmlOutput += `<li>${trimmedLine.replace(/^(-+\s*)/, "")}</li>`;
            prevLevel = level; // Update previous level
          }
        });

        // Close any remaining open lists
        while (stack.length > 0) {
          htmlOutput += `</ul>`;
          stack.pop();
        }

        // Update the cell with the final formatted HTML
        range.getCell(rowIndex + 1, colIndex + 1).setValue(htmlOutput);
      }
    });
  });
}

// Processes time calculation based on Status
const calculateTime = ({
  newStatus,
  prevStatus,
  cellStartTime,
  cellElapsedTime,
}) => {
  const now = new Date();
  if (newStatus === "In Progress") {
    return cellStartTime.setValue(now);
  } else if (
    ["Pending Entry", "On Hold", "Done", "Not Started", "Skipped"].includes(
      newStatus,
    ) &&
    prevStatus === "In Progress"
  ) {
    return setElapsedTime({
      startTime: new Date(cellStartTime.getValue()),
      endTime: now,
      cellElapsedTime,
    });
  }
};

// Calculates elapsed time between two dates
const setElapsedTime = ({ startTime, endTime, cellElapsedTime }) => {
  const prevElapsedTime = cellElapsedTime.getValue();
  const newElapsedTime = (endTime - startTime) / (1000 * 60); // Convert to minutes
  cellElapsedTime.setValue(prevElapsedTime + newElapsedTime);
};

const showExportPDFDialog = () => {
  const companyNames = getcompanyNames(); // Fetch project names
  const template = HtmlService.createTemplateFromFile("ExportForm");
  template.companyNames = companyNames; // Pass project names to the template
  const htmlOutput = template
    .evaluate()
    //.setWidth(650)
    .setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Export Time Entries to PDF",
  );
};

const getcompanyNames = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const namedRange = spreadsheet.getRangeByName("companyNames");
  const companyNames = namedRange
    .getValues()
    .flat()
    .filter((name) => name.trim() !== "");
  return companyNames;
};

const getCompanyProjects = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const namedRange = spreadsheet.getRangeByName("companyProjects");
  const data = namedRange.getValues();

  let companies = {};
  for (let i = 1; i < data.length; i++) {
    // Skip the header row
    const company = data[i][0]; // Assuming company is in column A
    const project = data[i][1]; // Assuming project is in column B

    if (!companies[company]) {
      companies[company] = [];
    }
    companies[company].push(project);
  }
  return companies;
};

const getProjectsForCompany = (company) => {
  const companyProjectData = getCompanyProjects();
  return companyProjectData[company] || [];
};

const createCompanyDropdown = () => {
  const companies = Object.keys(getCompanyProjects());
  const optionsHtml = companies
    .map((company) => `<option value="${company}">${company}</option>`)
    .join("");
  return optionsHtml;
};
