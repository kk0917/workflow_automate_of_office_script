function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  const name = 'queryAndDeleteFromCollection';

  // Get any worksheet with that name.
  const sheet = workbook.getWorksheet(name);

  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }

  // Add a blank worksheet with the name "queryAndDeleteFromCollection".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  const newSheet = workbook.addWorksheet(name);

  // Switch to the new worksheet.
  newSheet.activate();
}