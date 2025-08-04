// time-tracker/index.js

const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Export")
    .addItem("Export Time Entries to PDF", "showExportPDFDialog")
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
