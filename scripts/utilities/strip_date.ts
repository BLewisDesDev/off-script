/**
 * Utility Script: Date Homogenization for Entire Workbook
 *
 * This script finds all date-formatted cells with values in the active workbook,
 * converts them to a consistent DD/MM/YYYY text format, and removes date formatting.
 *
 * Purpose:
 * - Homogenize mixed date values (strings, serial numbers, formatted dates)
 * - Convert all to text format with consistent DD/MM/YYYY pattern
 * - Remove date formatting from cells to prevent future conversion issues
 */

function main(workbook: ExcelScript.Workbook): void {
  console.log('🚀 Starting Date Homogenization Utility...');

  try {
    const worksheets = workbook.getWorksheets();
    let totalProcessed = 0;
    const worksheetResults: string[] = [];

    for (let i = 0; i < worksheets.length; i++) {
      const worksheet = worksheets[i];
      console.log(
        `📋 Processing worksheet ${i + 1}/${worksheets.length}: ${worksheet.getName()}`
      );

      const processed = homogenizeDatesInWorksheet(worksheet);
      totalProcessed += processed;
      worksheetResults.push(`${worksheet.getName()}: ${processed} cells`);

      console.log(
        `✅ Completed worksheet: ${worksheet.getName()} (${processed} cells processed)`
      );
    }

    console.log('📋 Final Results:');
    for (const result of worksheetResults) {
      console.log(`  ${result}`);
    }

    console.log(
      `🎉 Date homogenization complete! Total cells processed: ${totalProcessed}`
    );
  } catch (error) {
    console.log('❌ Error during date homogenization:');
    if (error instanceof Error) {
      console.log(`Error message: ${error.message}`);
    }
  }
}

/**
 * Process a single worksheet to find and homogenize date cells
 */
function homogenizeDatesInWorksheet(worksheet: ExcelScript.Worksheet): number {
  const usedRange = worksheet.getUsedRange();
  if (!usedRange) {
    return 0;
  }

  // Get used range dimensions for calculations
  const values = usedRange.getValues();
  const rowCount = usedRange.getRowCount();
  const columnCount = usedRange.getColumnCount();
  const startRow = usedRange.getRowIndex();
  const startColumn = usedRange.getColumnIndex();

  let processedCount = 0;
  const updatedCells: { row: number; col: number; value: string }[] = [];
  const BATCH_SIZE = 200;
  let rowsProcessed = 0;

  console.log(
    `   📊 Scanning ${rowCount} rows × ${columnCount} columns (${rowCount * columnCount} cells)`
  );

  // Scan all cells for date values
  for (let row = 0; row < rowCount; row++) {
    for (let col = 0; col < columnCount; col++) {
      const cellValue = values[row][col];

      // Skip empty cells
      if (!cellValue) continue;

      // Check if this looks like a date value
      if (isDateValue(cellValue)) {
        const standardizedDate = standardizeDateValue(cellValue);
        if (standardizedDate) {
          updatedCells.push({
            row: startRow + row,
            col: startColumn + col,
            value: standardizedDate,
          });
          processedCount++;
        }
      }
    }

    rowsProcessed++;

    // Log progress every 200 rows
    if (rowsProcessed % BATCH_SIZE === 0) {
      console.log(
        `   ⏳ Processed ${rowsProcessed}/${rowCount} rows (${processedCount} date cells found so far)`
      );
    }
  }

  // Apply all updates in batches for efficiency
  if (updatedCells.length > 0) {
    console.log(
      `   🔄 Applying ${updatedCells.length} date standardizations...`
    );
    applyDateUpdates(worksheet, updatedCells);
  }

  return processedCount;
}

/**
 * Determine if a value appears to be a date
 */
function isDateValue(value: unknown): boolean {
  if (!value) return false;

  // Check for Excel serial numbers (typically dates are > 1000)
  if (typeof value === 'number' && value > 1000 && value < 100000) {
    return true;
  }

  // Check for date-like strings
  if (typeof value === 'string') {
    const str = value.toString().trim();

    // Common date patterns
    const datePatterns = [
      /^\d{1,2}\/\d{1,2}\/\d{4}$/, // DD/MM/YYYY or MM/DD/YYYY
      /^\d{4}-\d{1,2}-\d{1,2}$/, // YYYY-MM-DD
      /^\d{1,2}-\d{1,2}-\d{4}$/, // DD-MM-YYYY
      /^\d{1,2}\.\d{1,2}\.\d{4}$/, // DD.MM.YYYY
      /^\w{3}\s+\d{1,2},?\s+\d{4}$/, // Mon DD, YYYY
      /^\d{1,2}\s+\w{3}\s+\d{4}$/, // DD Mon YYYY
    ];

    return datePatterns.some(pattern => pattern.test(str));
  }

  return false;
}

/**
 * Convert various date formats to standardized DD/MM/YYYY text
 */
function standardizeDateValue(value: unknown): string | null {
  if (!value) return null;

  try {
    let date: Date;

    // Handle Excel serial numbers
    if (typeof value === 'number' && value > 1000) {
      // Convert Excel serial number to JavaScript Date
      const excelEpoch = new Date(1900, 0, 1);
      const daysSinceEpoch = value - 1;
      date = new Date(
        excelEpoch.getTime() + daysSinceEpoch * 24 * 60 * 60 * 1000
      );

      // Adjust for Excel's leap year bug
      if (value > 59) {
        date = new Date(date.getTime() - 24 * 60 * 60 * 1000);
      }
    } else {
      // Try to parse as date string
      date = new Date(value.toString());
    }

    // Validate the date
    if (isNaN(date.getTime())) {
      return null;
    }

    // Format as DD/MM/YYYY
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear().toString();

    return `${day}/${month}/${year}`;
  } catch {
    return null;
  }
}

/**
 * Apply date updates to worksheet in batches
 */
function applyDateUpdates(
  worksheet: ExcelScript.Worksheet,
  updates: { row: number; col: number; value: string }[]
): void {
  // Convert Map iteration to array-based approach for Office Scripts compatibility
  const updatesByRow: { [key: number]: { col: number; value: string }[] } = {};

  for (const update of updates) {
    if (!updatesByRow[update.row]) {
      updatesByRow[update.row] = [];
    }
    updatesByRow[update.row].push({ col: update.col, value: update.value });
  }

  // Apply updates row by row
  const rowKeys = Object.keys(updatesByRow).map(key => parseInt(key));
  for (const rowIndex of rowKeys) {
    const rowUpdates = updatesByRow[rowIndex];
    for (const cellUpdate of rowUpdates) {
      const cell = worksheet.getCell(rowIndex, cellUpdate.col);

      // Format as text first to prevent auto-conversion
      cell.setNumberFormat('@');

      // Set the standardized date value
      cell.setValue(cellUpdate.value);
    }
  }
}

/**
 * Alternative function: Process specific range
 * Usage: Call this instead of main() if you want to process only a selected range
 */
function processSelectedRange(workbook: ExcelScript.Workbook): void {
  console.log('🎯 Processing selected range...');

  const selectedRange = workbook.getSelectedRange();
  if (!selectedRange) {
    console.log('⚠️ No range selected');
    return;
  }

  const worksheet = selectedRange.getWorksheet();
  const values = selectedRange.getValues();
  const rowCount = selectedRange.getRowCount();
  const columnCount = selectedRange.getColumnCount();
  const startRow = selectedRange.getRowIndex();
  const startColumn = selectedRange.getColumnIndex();

  let processedCount = 0;
  const updatedCells: { row: number; col: number; value: string }[] = [];

  for (let row = 0; row < rowCount; row++) {
    for (let col = 0; col < columnCount; col++) {
      const cellValue = values[row][col];

      if (isDateValue(cellValue)) {
        const standardizedDate = standardizeDateValue(cellValue);
        if (standardizedDate) {
          updatedCells.push({
            row: startRow + row,
            col: startColumn + col,
            value: standardizedDate,
          });
          processedCount++;
        }
      }
    }
  }

  if (updatedCells.length > 0) {
    applyDateUpdates(worksheet, updatedCells);
  }

  console.log(`✅ Processed ${processedCount} date cells in selected range`);
}
