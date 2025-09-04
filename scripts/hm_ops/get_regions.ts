function main(workbook: ExcelScript.Workbook) {
  console.log('Starting postcode region lookup...');

  try {
    // Get worksheets
    const dataValuesSheet = workbook.getWorksheet('Values&Scripts');
    const ScheduleSheet = workbook.getWorksheet('Schedule');

    if (!dataValuesSheet) {
      console.log("Sheet 'Data Values' not found!");
      return;
    }

    if (!ScheduleSheet) {
      console.log("Sheet 'Schedule' not found!");
      return;
    }

    // Build fast lookup dictionary from Data Values sheet
    const postcodeToRegions = buildPostcodeLookup(dataValuesSheet);
    console.log(
      `Built lookup dictionary with ${postcodeToRegions.size} unique postcodes`
    );

    // Update Schedule with region matches
    updateClientRegions(ScheduleSheet, postcodeToRegions);

    console.log('Postcode region lookup completed successfully!');
  } catch (error) {
    console.log('Error during execution:', error);
  }
}

/**
 * Builds a fast lookup Map from postcodes to region arrays
 * Reads from columns E-K (regions) on Data Values sheet
 */
function buildPostcodeLookup(
  sheet: ExcelScript.Worksheet
): Map<string, string[]> {
  // Region columns E-K (indices 4-10)
  const regionColumns = ['E', 'F', 'G', 'H', 'I', 'J', 'K'];
  const regionNames = [
    'Western-Sydney',
    'South-West',
    'Inner-West',
    'South-East',
    'Northern-Sydney',
    'Hunter',
    'Illawarra',
  ];

  // Get the used range to determine how many rows have data
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log('No data found in Data Values sheet');
    return new Map();
  }

  const lastRow = usedRange.getRowCount();
  console.log(`Processing ${lastRow - 1} data rows (excluding header)`);

  // Create the lookup map
  const postcodeMap = new Map<string, string[]>();

  // Process each region column
  for (let colIndex = 0; colIndex < regionColumns.length; colIndex++) {
    const columnLetter = regionColumns[colIndex];
    const regionName = regionNames[colIndex];

    // Get all values in this column (starting from row 2 to skip header)
    const columnRange = sheet.getRange(
      `${columnLetter}2:${columnLetter}${lastRow}`
    );
    const columnValues = columnRange.getValues();

    // Process each postcode in this column
    for (let rowIndex = 0; rowIndex < columnValues.length; rowIndex++) {
      const cellValue = columnValues[rowIndex][0];

      // Convert to string and clean up
      const postcode = String(cellValue).trim();

      // Skip empty cells
      if (
        !postcode ||
        postcode === '' ||
        postcode === 'undefined' ||
        postcode === 'null'
      ) {
        continue;
      }

      // Add this region to the postcode's region list
      if (postcodeMap.has(postcode)) {
        const existingRegions = postcodeMap.get(postcode)!;
        if (!existingRegions.includes(regionName)) {
          existingRegions.push(regionName);
        }
      } else {
        postcodeMap.set(postcode, [regionName]);
      }
    }
  }

  return postcodeMap;
}

/**
 * Updates the Schedule sheet with matching regions for each postcode
 */
function updateClientRegions(
  sheet: ExcelScript.Worksheet,
  postcodeMap: Map<string, string[]>
) {
  // Find the header row to locate PostCode and Region columns
  const headerRange = sheet.getRange('A2:Z2'); // Check first row for headers
  const headerValues = headerRange.getValues()[0];

  let postcodeColumnIndex = -1;
  let regionColumnIndex = -1;

  // Find column indices
  for (let i = 0; i < headerValues.length; i++) {
    const header = String(headerValues[i]).toLowerCase().trim();
    if (header === 'postcode' || header === 'post code') {
      postcodeColumnIndex = i;
    }
    if (header === 'region' || header === 'regions') {
      regionColumnIndex = i;
    }
  }

  if (postcodeColumnIndex === -1) {
    console.log('PostCode column not found in Schedule sheet');
    return;
  }

  if (regionColumnIndex === -1) {
    console.log('Region column not found in Schedule sheet');
    return;
  }

  console.log(
    `Found PostCode in column ${String.fromCharCode(65 + postcodeColumnIndex)}, Region in column ${String.fromCharCode(65 + regionColumnIndex)}`
  );

  // Get the used range to find all client data
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log('No data found in Schedule sheet');
    return;
  }

  const lastRow = usedRange.getRowCount();
  console.log(`Processing ${lastRow - 1} client records`);

  // Get all postcode values (starting from row 3 to skip header)
  const postcodeColumn = String.fromCharCode(65 + postcodeColumnIndex);
  const postcodeRange = sheet.getRange(
    `${postcodeColumn}3:${postcodeColumn}${lastRow}`
  );
  const postcodeValues = postcodeRange.getValues();

  // Prepare region updates
  const regionUpdates: string[][] = [];
  let matchCount = 0;
  let noMatchCount = 0;

  for (let rowIndex = 0; rowIndex < postcodeValues.length; rowIndex++) {
    const postcode = String(postcodeValues[rowIndex][0]).trim();

    if (postcodeMap.has(postcode)) {
      const regions = postcodeMap.get(postcode)!;
      const regionString = regions.join(', ');
      regionUpdates.push([regionString]);
      matchCount++;
    } else {
      // No match found - you can change this to "No Match" if preferred
      regionUpdates.push(['']);
      console.log('No Region Match PostCode: ', postcode);
      noMatchCount++;
    }
  }

  // Update the Region column with all matches
  const regionColumn = String.fromCharCode(65 + regionColumnIndex);
  const regionRange = sheet.getRange(
    `${regionColumn}3:${regionColumn}${lastRow}`
  );
  regionRange.setValues(regionUpdates);

  console.log(`Updated ${matchCount} postcodes with region matches`);
  console.log(`${noMatchCount} postcodes had no matches`);
}
