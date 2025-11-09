// time-tracker/document-config.js

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
