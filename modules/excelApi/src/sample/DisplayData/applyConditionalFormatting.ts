function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  const range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  const conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom);

  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
