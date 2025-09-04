/**
 * Maps data from Cycle WK2 sheet to Schedule sheet
 * Matches on AC  NUMBER (Cycle WK2) to ACN (Schedule)
 * Only processes rows with AC  NUMBER values
 * Maps CREW to Team with conflict reporting
 * Converts DAY to "{DAY} - Week 2" format for Cycle
 */
function main(workbook: ExcelScript.Workbook) {
  // Get worksheets
  const cycleWk2Sheet = workbook.getWorksheet('Cycle WK2');
  const scheduleSheet = workbook.getWorksheet('Schedule');

  // Validate sheets exist
  if (!cycleWk2Sheet) {
    console.log("Error: 'Cycle WK2' sheet not found");
    return;
  }

  if (!scheduleSheet) {
    console.log("Error: 'Schedule' sheet not found");
    return;
  }

  // Get data from Cycle WK2 sheet
  const cycleWk2Range = cycleWk2Sheet.getUsedRange();
  if (!cycleWk2Range) {
    console.log('Error: No data found in Cycle WK2 sheet');
    return;
  }

  const cycleWk2Data = cycleWk2Range.getValues();
  const cycleWk2Headers = cycleWk2Data[1] as string[]; // Headers in row 2

  // Find column indices in Cycle WK2 sheet
  const acNumberColIndex = cycleWk2Headers.indexOf('AC  NUMBER'); // Note: 2 spaces
  const dayColIndex = cycleWk2Headers.indexOf('DAY');
  const crewColIndex = cycleWk2Headers.indexOf('CREW');
  const nameColIndex = cycleWk2Headers.indexOf('NAME');
  const suburbColIndex = cycleWk2Headers.indexOf('SUBURB');
  const postcodeColIndex = cycleWk2Headers.indexOf('POSTCODE');
  const phoneColIndex = cycleWk2Headers.indexOf('PHONE NUMBER');
  const paymentColIndex = cycleWk2Headers.indexOf('PAYMENT ');
  const orderColIndex = cycleWk2Headers.indexOf('ORDER '); // Note: trailing space
  const addressColIndex = cycleWk2Headers.indexOf('ADDRESS');
  const additionalNotesColIndex = cycleWk2Headers.indexOf('ADDITIONAL NOTES ');
  const complaintsColIndex = cycleWk2Headers.indexOf('Complaints/ Incidents '); // Note: different case and trailing space

  if (acNumberColIndex === -1) {
    console.log('Error: AC  NUMBER column not found in Cycle WK2 sheet');
    return;
  }

  // Log any missing source columns
  const missingSourceColumns: string[] = [];
  if (dayColIndex === -1) missingSourceColumns.push('DAY');
  if (crewColIndex === -1) missingSourceColumns.push('CREW');
  if (nameColIndex === -1) missingSourceColumns.push('NAME');
  if (suburbColIndex === -1) missingSourceColumns.push('SUBURB');
  if (postcodeColIndex === -1) missingSourceColumns.push('POSTCODE');
  if (phoneColIndex === -1) missingSourceColumns.push('PHONE NUMBER');
  if (paymentColIndex === -1) missingSourceColumns.push('PAYMENT');
  if (orderColIndex === -1) missingSourceColumns.push('ORDER ');
  if (addressColIndex === -1) missingSourceColumns.push('ADDRESS');
  if (additionalNotesColIndex === -1)
    missingSourceColumns.push('ADDITIONAL NOTES');
  if (complaintsColIndex === -1)
    missingSourceColumns.push('Complaints/ Incidents ');

  if (missingSourceColumns.length > 0) {
    console.log('Warning: Missing columns in Cycle WK2 sheet:');
    missingSourceColumns.forEach(col => console.log(`- ${col}`));
  }

  // Get data from Schedule sheet
  const scheduleRange = scheduleSheet.getUsedRange();
  if (!scheduleRange) {
    console.log('Error: No data found in Schedule sheet');
    return;
  }

  const scheduleData = scheduleRange.getValues();
  const scheduleHeaders = scheduleData[1] as string[];

  // Find column indices in Schedule sheet
  const acnColIndex = scheduleHeaders.indexOf('ACN');
  const firstNameColIndex = scheduleHeaders.indexOf('First Name');
  const lastNameColIndex = scheduleHeaders.indexOf('Last Name');
  const addressScheduleColIndex = scheduleHeaders.indexOf('Address');
  const suburbScheduleColIndex = scheduleHeaders.indexOf('Suburb');
  const postCodeColIndex = scheduleHeaders.indexOf('Post Code');
  const contactColIndex = scheduleHeaders.indexOf('Contact');
  const teamColIndex = scheduleHeaders.indexOf('Team');
  const orderScheduleColIndex = scheduleHeaders.indexOf('Order');
  const cycleColIndex = scheduleHeaders.indexOf('Cycle');
  const feeScheduleColIndex = scheduleHeaders.indexOf('Fee');
  const scheduleNotesColIndex = scheduleHeaders.indexOf('Notes');
  const complaintsScheduleColIndex = scheduleHeaders.indexOf(
    'Complaints/Incidents'
  );

  if (acnColIndex === -1) {
    console.log('Error: ACN column not found in Schedule sheet');
    return;
  }

  // Log any missing target columns
  const missingTargetColumns: string[] = [];
  if (firstNameColIndex === -1) missingTargetColumns.push('First Name');
  if (lastNameColIndex === -1) missingTargetColumns.push('Last Name');
  if (addressScheduleColIndex === -1) missingTargetColumns.push('Address');
  if (suburbScheduleColIndex === -1) missingTargetColumns.push('Suburb');
  if (postCodeColIndex === -1) missingTargetColumns.push('Post Code');
  if (contactColIndex === -1) missingTargetColumns.push('Contact');
  if (teamColIndex === -1) missingTargetColumns.push('Team');
  if (orderScheduleColIndex === -1) missingTargetColumns.push('Order');
  if (cycleColIndex === -1) missingTargetColumns.push('Cycle');
  if (feeScheduleColIndex === -1) missingTargetColumns.push('Fee');
  if (scheduleNotesColIndex === -1) missingTargetColumns.push('Notes');
  if (complaintsScheduleColIndex === -1)
    missingTargetColumns.push('Complaints/Incidents');

  if (missingTargetColumns.length > 0) {
    console.log('Warning: Missing columns in Schedule sheet:');
    missingTargetColumns.forEach(col => console.log(`- ${col}`));
  }

  // Create cycle lookup table for fast matching
  const cycleLookupTable = createCycleLookupTable();

  // Create lookup map from Cycle WK2 data - data starts from row 3 (index 2)
  const cycleWk2Lookup = new Map<
    string,
    {
      day: string;
      crew: string;
      name: string;
      suburb: string;
      postcode: string;
      phone: string;
      payment: string;
      order: string;
      address: string;
      additionalNotes: string;
      complaints: string;
    }
  >();

  for (let i = 2; i < cycleWk2Data.length; i++) {
    // Start from row 3 (index 2)
    const acNumber = cycleWk2Data[i][acNumberColIndex];
    // Check if AC NUMBER exists
    if (!acNumber) {
      // console.log(`Row ${i + 1}: Missing AC NUMBER`);
      continue;
    }

    const trimmedAcNumber = String(acNumber).trim();

    // Check if AC NUMBER is empty after trimming
    if (trimmedAcNumber === '') {
      // console.log(`Row ${i + 1}: Empty AC NUMBER (whitespace only)`);
      continue;
    }

    // Validate AC NUMBER format
    if (!/^AC\d{8}$/.test(trimmedAcNumber)) {
      console.log(
        `Row ${i + 1}: Malformed AC NUMBER: "${trimmedAcNumber}" (Expected: AC + 8 digits)`
      );
      continue;
    }

    cycleWk2Lookup.set(trimmedAcNumber, {
      day:
        dayColIndex !== -1 && cycleWk2Data[i][dayColIndex]
          ? String(cycleWk2Data[i][dayColIndex]).trim()
          : '',
      crew:
        crewColIndex !== -1 && cycleWk2Data[i][crewColIndex]
          ? String(cycleWk2Data[i][crewColIndex]).trim()
          : '',
      name:
        nameColIndex !== -1 && cycleWk2Data[i][nameColIndex]
          ? String(cycleWk2Data[i][nameColIndex]).trim()
          : '',
      suburb:
        suburbColIndex !== -1 && cycleWk2Data[i][suburbColIndex]
          ? String(cycleWk2Data[i][suburbColIndex]).trim()
          : '',
      postcode:
        postcodeColIndex !== -1 && cycleWk2Data[i][postcodeColIndex]
          ? String(cycleWk2Data[i][postcodeColIndex]).trim()
          : '',
      phone:
        phoneColIndex !== -1 && cycleWk2Data[i][phoneColIndex]
          ? String(cycleWk2Data[i][phoneColIndex]).trim()
          : '',
      payment:
        paymentColIndex !== -1 && cycleWk2Data[i][paymentColIndex]
          ? String(cycleWk2Data[i][paymentColIndex]).trim()
          : '',
      order:
        orderColIndex !== -1 && cycleWk2Data[i][orderColIndex]
          ? String(cycleWk2Data[i][orderColIndex]).trim()
          : '',
      address:
        addressColIndex !== -1 && cycleWk2Data[i][addressColIndex]
          ? String(cycleWk2Data[i][addressColIndex]).trim()
          : '',
      additionalNotes:
        additionalNotesColIndex !== -1 &&
        cycleWk2Data[i][additionalNotesColIndex]
          ? String(cycleWk2Data[i][additionalNotesColIndex]).trim()
          : '',
      complaints:
        complaintsColIndex !== -1 && cycleWk2Data[i][complaintsColIndex]
          ? String(cycleWk2Data[i][complaintsColIndex]).trim()
          : '',
    });
  }

  // Update matching records in Schedule sheet
  let updatedCount = 0;
  let matchedCount = 0;
  let teamConflicts = 0;
  const matchedCycleWk2Records = new Set<string>();
  const teamConflictReports: string[] = [];

  for (let i = 1; i < scheduleData.length; i++) {
    // Schedule data from row 2 (index 1)
    const acn = scheduleData[i][acnColIndex];

    if (acn && String(acn).trim() !== '') {
      const acnString = String(acn).trim();
      const cycleRecord = cycleWk2Lookup.get(acnString);

      if (cycleRecord) {
        matchedCount++;
        matchedCycleWk2Records.add(acnString);

        // Parse name into first and last name
        const fullName = cycleRecord.name.trim();
        const nameParts = fullName.split(' ').filter(part => part.length > 0);
        const firstName =
          nameParts.length > 1 ? nameParts.slice(0, -1).join(' ') : fullName;
        const lastName =
          nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';

        // Convert DAY to proper cycle format
        const cycle = formatDayToCycle(cycleRecord.day, cycleLookupTable);

        // Parse fee from payment
        const fee = parsePaymentToFee(cycleRecord.payment);

        // Check Team field for conflicts first
        if (teamColIndex !== -1) {
          const currentTeam = scheduleData[i][teamColIndex];
          const currentTeamStr = currentTeam ? String(currentTeam).trim() : '';

          if (currentTeamStr !== '' && currentTeamStr !== cycleRecord.crew) {
            // Team conflict found - skip this entire row
            teamConflicts++;
            teamConflictReports.push(
              `ACN ${acnString}: Team is "${currentTeamStr}", expected "${cycleRecord.crew}"`
            );
            continue; // Skip to next record
          }

          // Update Team if empty
          if (currentTeamStr === '') {
            scheduleSheet.getCell(i, teamColIndex).setValue(cycleRecord.crew);
            updatedCount++;
          }
        }

        // Update all other fields if empty/whitespace (only reached if no team conflict)
        if (
          firstNameColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][firstNameColIndex])
        ) {
          scheduleSheet.getCell(i, firstNameColIndex).setValue(firstName);
          updatedCount++;
        }
        if (
          lastNameColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][lastNameColIndex])
        ) {
          scheduleSheet.getCell(i, lastNameColIndex).setValue(lastName);
          updatedCount++;
        }
        if (
          addressScheduleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][addressScheduleColIndex])
        ) {
          scheduleSheet
            .getCell(i, addressScheduleColIndex)
            .setValue(cycleRecord.address);
          updatedCount++;
        }
        if (
          suburbScheduleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][suburbScheduleColIndex])
        ) {
          scheduleSheet
            .getCell(i, suburbScheduleColIndex)
            .setValue(cycleRecord.suburb);
          updatedCount++;
        }
        if (
          postCodeColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][postCodeColIndex])
        ) {
          scheduleSheet
            .getCell(i, postCodeColIndex)
            .setValue(cycleRecord.postcode);
          updatedCount++;
        }
        if (
          contactColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][contactColIndex])
        ) {
          scheduleSheet.getCell(i, contactColIndex).setValue(cycleRecord.phone);
          updatedCount++;
        }
        if (
          orderScheduleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][orderScheduleColIndex])
        ) {
          scheduleSheet
            .getCell(i, orderScheduleColIndex)
            .setValue(cycleRecord.order);
          updatedCount++;
        }
        if (
          cycleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][cycleColIndex])
        ) {
          scheduleSheet.getCell(i, cycleColIndex).setValue(cycle);
          updatedCount++;
        }
        if (
          feeScheduleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][feeScheduleColIndex])
        ) {
          scheduleSheet.getCell(i, feeScheduleColIndex).setValue(fee);
          updatedCount++;
        }
        if (
          scheduleNotesColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][scheduleNotesColIndex])
        ) {
          scheduleSheet
            .getCell(i, scheduleNotesColIndex)
            .setValue(cycleRecord.additionalNotes);
          updatedCount++;
        }
        if (
          complaintsScheduleColIndex !== -1 &&
          isEmptyOrWhitespace(scheduleData[i][complaintsScheduleColIndex])
        ) {
          scheduleSheet
            .getCell(i, complaintsScheduleColIndex)
            .setValue(cycleRecord.complaints);
          updatedCount++;
        }
      }
    }
  }

  // Find Cycle WK2 records that didn't match any Schedule record
  const unmatchedCycleWk2Records: string[] = [];
  for (const [acNumber] of cycleWk2Lookup) {
    if (!matchedCycleWk2Records.has(acNumber)) {
      unmatchedCycleWk2Records.push(acNumber);
    }
  }

  // Add unmatched Cycle WK2 records to Schedule sheet
  const addedCount = addUnmatchedCycleWk2Records(
    scheduleSheet,
    scheduleHeaders,
    unmatchedCycleWk2Records,
    cycleWk2Lookup,
    cycleLookupTable
  );

  // Report results
  console.log(`Processing complete:`);
  console.log(`- Cycle WK2 records processed: ${cycleWk2Lookup.size}`);
  console.log(`- Cycle WK2 records matched: ${matchedCycleWk2Records.size}`);
  console.log(
    `- Cycle WK2 records unmatched: ${unmatchedCycleWk2Records.length}`
  );
  console.log(`- Schedule records matched: ${matchedCount}`);
  console.log(`- Cells updated: ${updatedCount}`);
  console.log(`- New records added: ${addedCount}`);
  console.log(`- Team conflicts: ${teamConflicts}`);

  // Report team conflicts
  if (teamConflictReports.length > 0) {
    console.log('Team conflicts found:');
    teamConflictReports.forEach(conflict => console.log(`- ${conflict}`));
  }
}

/**
 * Creates lookup table for valid cycle values for fast matching
 */
function createCycleLookupTable(): string[] {
  const days = [
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
    'Sunday',
  ];
  const weeks = ['Week 1', 'Week 2', 'Week 3', 'Week 4'];
  const cycles: string[] = [];

  for (const day of days) {
    for (const week of weeks) {
      cycles.push(`${day} - ${week}`);
    }
  }

  return cycles;
}

/**
 * Formats DAY field to "{DAY} - Week 2" format
 */
function formatDayToCycle(day: string, cycleLookupTable: string[]): string {
  const trimmedDay = day.trim();

  // Create the target format
  const targetCycle = `${trimmedDay} - Week 2`;

  // Try exact match first
  if (cycleLookupTable.includes(targetCycle)) {
    return targetCycle;
  }

  // Fast approximate matching - find best score
  let bestMatch = cycleLookupTable[0];
  let bestScore = 0;

  for (const cycleValue of cycleLookupTable) {
    const score = calculateSimpleMatchScore(
      targetCycle.toLowerCase(),
      cycleValue.toLowerCase()
    );
    if (score > bestScore) {
      bestScore = score;
      bestMatch = cycleValue;
    }
  }

  return bestMatch;
}

/**
 * Fast simple string matching - counts matching characters
 */
function calculateSimpleMatchScore(str1: string, str2: string): number {
  let score = 0;
  const minLength = Math.min(str1.length, str2.length);

  for (let i = 0; i < minLength; i++) {
    if (str1[i] === str2[i]) {
      score++;
    }
  }

  return score;
}

/**
 * Parses payment string and extracts numeric value as float
 * Returns 0.0 if no numbers found but string is not null/empty
 */
function parsePaymentToFee(payment: string): number {
  if (!payment || payment.trim() === '') {
    return 0.0;
  }

  // Extract all numbers from the string
  const numberMatch = payment.match(/\d+\.?\d*/);

  if (numberMatch) {
    const parsedValue = parseFloat(numberMatch[0]);
    return isNaN(parsedValue) ? 0.0 : parsedValue;
  }

  return 0.0;
}

/**
 * Checks if a value is empty or contains only whitespace
 */
function isEmptyOrWhitespace(
  value: string | number | boolean | null | undefined
): boolean {
  return !value || String(value).trim() === '';
}

/**
 * Adds unmatched Cycle WK2 records as new rows in Schedule sheet
 */
function addUnmatchedCycleWk2Records(
  scheduleSheet: ExcelScript.Worksheet,
  scheduleHeaders: string[],
  unmatchedACNumbers: string[],
  cycleWk2Lookup: Map<
    string,
    {
      day: string;
      crew: string;
      name: string;
      suburb: string;
      postcode: string;
      phone: string;
      payment: string;
      order: string;
      address: string;
      additionalNotes: string;
      complaints: string;
    }
  >,
  cycleLookupTable: string[]
): number {
  if (unmatchedACNumbers.length === 0) {
    return 0;
  }

  // Find column indices in Schedule sheet
  const acnColIndex: number = scheduleHeaders.indexOf('ACN');
  const firstNameColIndex: number = scheduleHeaders.indexOf('First Name');
  const lastNameColIndex: number = scheduleHeaders.indexOf('Last Name');
  const addressColIndex: number = scheduleHeaders.indexOf('Address');
  const suburbColIndex: number = scheduleHeaders.indexOf('Suburb');
  const postCodeColIndex: number = scheduleHeaders.indexOf('Post Code');
  const contactColIndex: number = scheduleHeaders.indexOf('Contact');
  const teamColIndex: number = scheduleHeaders.indexOf('Team');
  const orderColIndex: number = scheduleHeaders.indexOf('Order');
  const cycleColIndex: number = scheduleHeaders.indexOf('Cycle');
  const feeColIndex: number = scheduleHeaders.indexOf('Fee');
  const notesColIndex: number = scheduleHeaders.indexOf('Notes');
  const complaintsColIndex: number = scheduleHeaders.indexOf(
    'Complaints/Incidents'
  );

  // Get current used range to find next available row
  const usedRange = scheduleSheet.getUsedRange();
  let nextRow = usedRange ? usedRange.getRowCount() : 1;

  let addedCount = 0;

  for (const acNumber of unmatchedACNumbers) {
    const cycleRecord = cycleWk2Lookup.get(acNumber);
    if (!cycleRecord) continue;

    // Parse name into first and last name
    const fullName = cycleRecord.name.trim();
    const nameParts = fullName.split(' ').filter(part => part.length > 0);
    const firstName =
      nameParts.length > 1 ? nameParts.slice(0, -1).join(' ') : fullName;
    const lastName =
      nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';

    // Convert DAY to proper cycle format
    const cycle = formatDayToCycle(cycleRecord.day, cycleLookupTable);

    // Parse fee from payment
    const fee = parsePaymentToFee(cycleRecord.payment);

    // Set values for each column
    if (acnColIndex !== -1) {
      scheduleSheet.getCell(nextRow, acnColIndex).setValue(acNumber);
    }
    if (firstNameColIndex !== -1) {
      scheduleSheet.getCell(nextRow, firstNameColIndex).setValue(firstName);
    }
    if (lastNameColIndex !== -1) {
      scheduleSheet.getCell(nextRow, lastNameColIndex).setValue(lastName);
    }
    if (addressColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, addressColIndex)
        .setValue(cycleRecord.address);
    }
    if (suburbColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, suburbColIndex)
        .setValue(cycleRecord.suburb);
    }
    if (postCodeColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, postCodeColIndex)
        .setValue(cycleRecord.postcode);
    }
    if (contactColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, contactColIndex)
        .setValue(cycleRecord.phone);
    }
    if (teamColIndex !== -1) {
      scheduleSheet.getCell(nextRow, teamColIndex).setValue(cycleRecord.crew);
    }
    if (orderColIndex !== -1) {
      scheduleSheet.getCell(nextRow, orderColIndex).setValue(cycleRecord.order);
    }
    if (cycleColIndex !== -1) {
      scheduleSheet.getCell(nextRow, cycleColIndex).setValue(cycle);
    }
    if (feeColIndex !== -1) {
      scheduleSheet.getCell(nextRow, feeColIndex).setValue(fee);
    }
    if (notesColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, notesColIndex)
        .setValue(cycleRecord.additionalNotes);
    }
    if (complaintsColIndex !== -1) {
      scheduleSheet
        .getCell(nextRow, complaintsColIndex)
        .setValue(cycleRecord.complaints);
    }

    nextRow++;
    addedCount++;
  }

  return addedCount;
}
