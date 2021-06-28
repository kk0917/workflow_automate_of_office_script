function main(
    workbook: ExcelScript.Workbook,
    from: string,
    dateReceived: string,
    subject: string) {

    // Get the email table.
    const emailWorksheet = workbook.getWorksheet("Emails");
    const table = emailWorksheet.getTable("EmailTable");

    // Get the PivotTable.
    const pivotTableWorksheet = workbook.getWorksheet("Subjects");
    const pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

    //...
}