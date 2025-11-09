// time-tracker/sheet-data.js

const getCompanyNames = () => {
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
