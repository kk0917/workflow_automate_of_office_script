function main(workbook: ExcelScript.Workbook) {
  // Get the current used range.
  const range = workbook.getActiveWorksheet().getUsedRange();
  
  // Get all the blank cells.
  const blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

  // Highlight the blank cells with a yellow background.
  blankCells.getFormat().getFill().setColor("yellow");
}
