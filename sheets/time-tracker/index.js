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
