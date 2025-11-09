// time-tracker/data.js

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
