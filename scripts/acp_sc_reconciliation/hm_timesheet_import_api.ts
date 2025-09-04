/**
 * Office Script to sync ShiftCare timesheet data to spreadsheet
 * Fetches timesheets for a specific month and adds all data
 * Replace the placeholder constants with your actual credentials
 */

const SHIFTCARE_BASE_URL = 'https://api.shiftcare.com';
const ACCOUNT_ID = '653156';
const API_KEY = 'ZGYzOWFlZTIxYmEwNTdlYzFhOTNhNDIxZWRhOWExMzY2NTMxNTY=';
const MAX_ITEMS_PER_PAGE = 20;

// Date range configuration - update TARGET_MONTH to change the month
const DATE_RANGES = {
  january: '2025-01-01T00:00:00Z&to=2025-01-31T00:00:00Z',
  february: '2025-02-01T00:00:00Z&to=2025-02-28T00:00:00Z',
  march: '2025-03-01T00:00:00Z&to=2025-03-31T00:00:00Z',
  april: '2025-04-01T00:00:00Z&to=2025-04-30T00:00:00Z',
  may: '2025-05-01T00:00:00Z&to=2025-05-31T00:00:00Z',
  june: '2025-06-01T00:00:00Z&to=2025-06-30T00:00:00Z',
  july: '2025-07-01T00:00:00Z&to=2025-07-31T00:00:00Z',
  august: '2025-08-01T00:00:00Z&to=2025-08-31T00:00:00Z',
  september: '2025-09-01T00:00:00Z&to=2025-09-30T00:00:00Z',
  october: '2025-10-01T00:00:00Z&to=2025-10-31T00:00:00Z',
  november: '2025-11-01T00:00:00Z&to=2025-11-30T00:00:00Z',
  december: '2025-12-01T00:00:00Z&to=2025-12-31T00:00:00Z',
};

// **UPDATE THIS TO TARGET DIFFERENT MONTHS**
const TARGET_MONTH: keyof typeof DATE_RANGES = 'september';

interface TimesheetItem {
  payable_id: string;
  payable_type: string;
  payable_name: string;
  payable_unit?: string | null;
  start_at: string;
  finish_at: string;
  break_minutes: number;
  amount: number;
}

interface ShiftCareTimesheet {
  staff_id: string;
  date: string;
  client_ids: string[];
  items: TimesheetItem[];
  status: string;
}

interface ShiftCareMetadata {
  current_page: number | string;
  total_pages: number | string;
  total_items?: number | string;
  total_count?: number | string;
  per_page?: number | string;
  page_items?: number | string;
  first_page_link?: string;
  next_page_link?: string;
  previous_page_link?: string;
  last_page_link?: string;
}

interface ShiftCareTimesheetsResponse {
  timesheets: ShiftCareTimesheet[];
  _metadata: ShiftCareMetadata;
}

interface ProcessingStats {
  pagesProcessed: number;
  totalTimesheetsReviewed: number;
  totalItemsProcessed: number;
  itemsAdded: number;
  errors: number;
}

/**
 * Parses ISO date string to DD/MM/YYYY format
 * @param isoDateString - ISO date string like "2025-01-01T21:00:00Z"
 * @returns Formatted date string in DD/MM/YYYY format
 */
function parseIsoToDateString(isoDateString: string): string {
  try {
    if (!isoDateString || isoDateString.trim() === '') {
      return '';
    }

    const date = new Date(isoDateString);

    if (isNaN(date.getTime())) {
      console.log(`⚠️ Invalid date format: ${isoDateString}`);
      return isoDateString;
    }

    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear().toString();

    return `${day}/${month}/${year}`;
  } catch (error) {
    console.log(`⚠️ Error parsing date ${isoDateString}:`, error);
    return isoDateString;
  }
}

/**
 * Parses ISO date string to DD/MM/YYYY HH:MM format
 * @param isoDateString - ISO date string like "2025-01-01T21:00:00Z"
 * @returns Formatted datetime string in DD/MM/YYYY HH:MM format
 */
function parseIsoToDateTimeString(isoDateString: string): string {
  try {
    if (!isoDateString || isoDateString.trim() === '') {
      return '';
    }

    const date = new Date(isoDateString);

    if (isNaN(date.getTime())) {
      console.log(`⚠️ Invalid datetime format: ${isoDateString}`);
      return isoDateString;
    }

    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear().toString();
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');

    return `${day}/${month}/${year} ${hours}:${minutes}`;
  } catch (error) {
    console.log(`⚠️ Error parsing datetime ${isoDateString}:`, error);
    return isoDateString;
  }
}

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  console.log(`🚀 Starting ShiftCare timesheet sync for ${TARGET_MONTH}...`);

  const stats: ProcessingStats = {
    pagesProcessed: 0,
    totalTimesheetsReviewed: 0,
    totalItemsProcessed: 0,
    itemsAdded: 0,
    errors: 0,
  };

  try {
    // Get the target date range
    const dateRange = DATE_RANGES[TARGET_MONTH];
    if (!dateRange) {
      console.log(`❌ Invalid target month: ${TARGET_MONTH}`);
      return;
    }

    console.log(`📅 Target date range: ${dateRange}`);

    // Setup worksheet headers if needed
    const nextRow = setupWorksheetHeaders(workbook);
    if (nextRow === -1) {
      console.log('❌ Failed to setup worksheet');
      return;
    }

    // Collect all timesheet items
    let currentPage = 1;
    let hasMorePages = true;
    let currentRowPointer = nextRow;

    while (hasMorePages) {
      console.log(`\n📄 Processing page ${currentPage}...`);

      const pageResult = await fetchTimesheetsPage(currentPage, dateRange);
      if (!pageResult) {
        console.log(`❌ Failed to fetch page ${currentPage}`);
        stats.errors++;
        break;
      }

      stats.pagesProcessed++;
      stats.totalTimesheetsReviewed += pageResult.timesheets.length;

      // Process timesheets and add to worksheet immediately
      if (pageResult.timesheets.length > 0) {
        const addedCount = batchAddTimesheetItemsToWorksheet(
          workbook,
          pageResult.timesheets,
          currentRowPointer
        );
        stats.itemsAdded += addedCount;

        // Update row pointer for next batch
        currentRowPointer += addedCount;

        // Count total items processed
        const itemsOnPage = pageResult.timesheets.reduce(
          (total, timesheet) => total + timesheet.items.length,
          0
        );
        stats.totalItemsProcessed += itemsOnPage;
      }

      console.log(
        `📋 Page ${currentPage}: ${pageResult.timesheets.length} timesheets processed`
      );

      // Check if there are more pages
      if (pageResult.timesheets.length === 0) {
        console.log('🛑 No more timesheets found');
        hasMorePages = false;
      } else {
        currentPage++;
      }

      // Safety check - don't run forever
      if (currentPage > 500) {
        console.log('⚠️ Safety limit reached (500 pages). Stopping sync.');
        break;
      }
    }

    // Print final statistics
    printFinalStats(stats);
  } catch (error) {
    console.log('❌ Error during ShiftCare timesheet sync:');
    if (error instanceof Error) {
      console.log(`Error message: ${error.message}`);
    } else {
      console.log('Unknown error occurred');
    }
    stats.errors++;
    printFinalStats(stats);
  }
}

/**
 * Sets up the worksheet with headers for timesheet items and returns next available row
 */
function setupWorksheetHeaders(workbook: ExcelScript.Workbook): number {
  try {
    const worksheet = workbook.getWorksheet('HM-Sessions');
    if (!worksheet) {
      console.log("❌ Worksheet 'hm-sessions' not found!");
      return -1;
    }

    // Get existing used range to determine where to start
    const usedRange = worksheet.getUsedRange();
    let nextRow = 1;

    if (usedRange) {
      nextRow = usedRange.getRowCount() + 1;
      console.log(`📊 Found existing data. Starting at row ${nextRow}`);
      return nextRow;
    }

    // If no existing data, add headers
    const headers = [
      'Staff ID', // Column A
      'Timesheet Date', // Column B
      'Client IDs', // Column C
      'Status', // Column D
      'Payable ID', // Column E
      'Payable Type', // Column F
      'Payable Name', // Column G
      'Payable Unit', // Column H
      'Start Time', // Column I
      'Finish Time', // Column J
      'Break Minutes', // Column K
      'Amount', // Column L
    ];

    // Set headers in row 1
    const headerRange = worksheet.getRange(`A1:L1`);
    headerRange.setValues([headers]);

    // Format headers (bold)
    headerRange.getFormat().getFont().setBold(true);

    console.log('✅ Worksheet headers set up successfully');
    return 2; // Data starts from row 2
  } catch (error) {
    console.log('❌ Error setting up worksheet headers:');
    if (error instanceof Error) {
      console.log(`Error message: ${error.message}`);
    }
    return -1;
  }
}

/**
 * Fetches a single page of timesheets from the ShiftCare API
 */
async function fetchTimesheetsPage(
  page: number,
  dateRange: string
): Promise<ShiftCareTimesheetsResponse | null> {
  try {
    const apiUrl = `${SHIFTCARE_BASE_URL}/api/v3/timesheets?from=${dateRange}&page=${page}&per_page=${MAX_ITEMS_PER_PAGE}&include_metadata=true`;

    const credentials = btoa(`${ACCOUNT_ID}:${API_KEY}`);
    const headers = {
      Authorization: `Basic ${credentials}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    };

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: headers,
    });

    if (!response.ok) {
      throw new Error(
        `API request failed: ${response.status} ${response.statusText}`
      );
    }

    return await response.json();
  } catch (error) {
    console.log(`❌ Error fetching page ${page}:`);
    if (error instanceof Error) {
      console.log(`Error message: ${error.message}`);
    }
    return null;
  }
}

/**
 * Batch adds timesheet items to the worksheet starting from specified row
 */
function batchAddTimesheetItemsToWorksheet(
  workbook: ExcelScript.Workbook,
  timesheets: ShiftCareTimesheet[],
  startRow: number
): number {
  try {
    const worksheet = workbook.getWorksheet('hm-sessions');
    if (!worksheet) {
      console.log('❌ Worksheet not found!');
      return 0;
    }

    // Flatten timesheet items into rows
    const batchData: (string | number)[][] = [];

    for (const timesheet of timesheets) {
      // Join client IDs into a comma-separated string
      const clientIdsString = timesheet.client_ids.join(', ');

      // Format timesheet date
      const formattedDate = parseIsoToDateString(timesheet.date);

      // Process each item in the timesheet
      for (const item of timesheet.items) {
        const formattedStartTime = parseIsoToDateTimeString(item.start_at);
        const formattedFinishTime = parseIsoToDateTimeString(item.finish_at);

        const itemRow = [
          timesheet.staff_id, // Column A: Staff ID
          formattedDate, // Column B: Timesheet Date
          clientIdsString, // Column C: Client IDs
          timesheet.status, // Column D: Status
          item.payable_id, // Column E: Payable ID
          item.payable_type, // Column F: Payable Type
          item.payable_name, // Column G: Payable Name
          item.payable_unit || '', // Column H: Payable Unit
          formattedStartTime, // Column I: Start Time
          formattedFinishTime, // Column J: Finish Time
          item.break_minutes, // Column K: Break Minutes
          item.amount, // Column L: Amount
        ];

        batchData.push(itemRow);
      }
    }

    // Write data if we have any
    if (batchData.length > 0) {
      const endRow = startRow + batchData.length - 1;
      const batchRange = worksheet.getRange(`A${startRow}:L${endRow}`);
      batchRange.setValues(batchData);
      console.log(
        `✅ Added ${batchData.length} timesheet items to worksheet (rows ${startRow}-${endRow})`
      );
      return batchData.length;
    }

    return 0;
  } catch (error) {
    console.log('❌ Error adding timesheet items to worksheet:');
    if (error instanceof Error) {
      console.log(`Error message: ${error.message}`);
    }
    return 0;
  }
}

/**
 * Prints final processing statistics
 */
function printFinalStats(stats: ProcessingStats): void {
  console.log('\n' + '='.repeat(60));
  console.log('📊 FINAL PROCESSING STATISTICS');
  console.log('='.repeat(60));
  console.log(`📅 Target month: ${TARGET_MONTH}`);
  console.log(`📄 Pages processed: ${stats.pagesProcessed}`);
  console.log(`📋 Total timesheets reviewed: ${stats.totalTimesheetsReviewed}`);
  console.log(
    `🔢 Total timesheet items processed: ${stats.totalItemsProcessed}`
  );
  console.log(`📝 Items added to worksheet: ${stats.itemsAdded}`);
  console.log(`❌ Errors encountered: ${stats.errors}`);
  console.log('='.repeat(60));

  if (stats.itemsAdded > 0) {
    console.log('🎉 Timesheet sync completed successfully!');
  } else {
    console.log('❌ No timesheet items were added');
  }
}
