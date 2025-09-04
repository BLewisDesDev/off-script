/**
 * Office Script: Highlight High Values
 * Description: Highlights cells above a threshold with a specified color
 * Author: Development Environment
 * Created: $(date +%Y-%m-%d)
 */

function main(
  workbook: ExcelScript.Workbook,
  highlightThreshold: number = 100,
  color: string = 'yellow'
): void {
  // Get the active worksheet
  const worksheet = workbook.getActiveWorksheet();

  if (!worksheet) {
    console.log('No active worksheet found');
    return;
  }

  // Get the used range
  const usedRange = worksheet.getUsedRange();

  if (!usedRange) {
    console.log('No data found in worksheet');
    return;
  }

  // Get all values in the used range
  const values = usedRange.getValues();

  // Process each cell
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      const cellValue = values[row][col];

      // Check if value is a number and above threshold
      if (typeof cellValue === 'number' && cellValue > highlightThreshold) {
        // Get the specific cell and highlight it
        const cell = usedRange.getCell(row, col);
        cell.getFormat().getFill().setColor(color);
      }
    }
  }

  console.log(`Highlighted cells with values > ${highlightThreshold}`);
}
