function main(workbook: ExcelScript.Workbook) {
  // Get the current used range.
  let range = workbook.getActiveWorksheet().getUsedRange();
  
  // Get all the blank cells.
  let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

  // Highlight the blank cells with a yellow background.
  blankCells.getFormat().getFill().setColor("yellow");
}