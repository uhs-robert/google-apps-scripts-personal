// time-tracker/dialog.js

const createCompanyDropdown = () => {
  const companies = Object.keys(getCompanyProjects());
  const optionsHtml = companies
    .map((company) => `<option value="${company}">${company}</option>`)
    .join("");
  return optionsHtml;
};

const showExportPDFDialog = () => {
  const companyNames = getCompanyNames(); // Fetch project names
  const template = HtmlService.createTemplateFromFile("ExportForm");
  template.companyNames = companyNames;
  const htmlOutput = template
    .evaluate()
    //.setWidth(650)
    .setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Export Time Entries to PDF",
  );
};
