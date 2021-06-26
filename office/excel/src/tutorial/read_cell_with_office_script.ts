function readCell(workbook: ExcelScript.Workbook) {
    // Get the worksheet from the workbook.
    const worksheet = workbook.getWorksheet("call_from_the_power_automate");

    // Get the cells at A1 and B1.
    const dateRange = worksheet.getRange("A1");
    const timeRange = worksheet.getRange("B1");

    // Get the current date and time using the JavaScript Date object.
    const date = new Date(Date.now());

    // Add the date string to A1.
    dateRange.setValue(date.toLocaleDateString());

    // Add the time string to B1.
    timeRange.setValue(date.toLocaleTimeString());
}
