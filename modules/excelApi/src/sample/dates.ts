function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  const dateRange = workbook.getActiveWorksheet().getRange("A1");
  const timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  const date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}