function readRow(workbook: ExcelScript.Workbook) {
    // Get the current worksheet.
    let selectedSheet = workbook.getActiveWorksheet();

    /** Format the range to display numerical dollar amounts.
     * 
     * TODO:
     * The format set by the following code is not applied,
     *   but ¥ symbol and the other letters is applied instead of $.
     */
    selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

    // Fit the width of all the used columns to the data.
    selectedSheet.getUsedRange().getFormat().autofitColumns();

    // Get the values of the used range.
    let range       = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();

    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
}
