function main(workbook: ExcelScript.Workbook) : string {
  // Get the H1 worksheet.
  const worksheet = workbook.getWorksheet("H1");

  // Get the first (and only) table in the worksheet.
  const table = worksheet.getTables()[0];

  // Get the data from the table.
  const tableValues = table.getRangeBetweenHeaderAndTotal().getValues();

  // Look for the first row where today's date is between the row's start and end dates.
  const currentDate = new Date();

  for (let row = 0; row < tableValues.length; row++) {
    const startDate: Date = convertDate(tableValues[row][2] as number);
    const endDate:   Date = convertDate(tableValues[row][3] as number);

    if (startDate <= currentDate && endDate >= currentDate) {
      // Return the first matching email address.
      return tableValues[row][1].toString();
    }
  }
}

/** Convert the Excel date to a JavaScript Date object
 * 
 * @param   number excelDateValue 
 * @returns Date obj that converted js date from excel date
 */
function convertDate(excelDateValue: number): Date {
  const jsDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  return jsDate;
}
