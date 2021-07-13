function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  const sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  const names = sheets.map (sheet => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    const colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
