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

  // Parse the received date string to determine the day of the week.
  const emailDate = new Date(dateReceived);
  const dayName   = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}