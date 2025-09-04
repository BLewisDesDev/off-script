function main(workbook: ExcelScript.Workbook) {
  const sourceSheetName = 'Schedule';
  const targetSheetName = 'Cycle Week 1';

  // ---- Config ----
  const dayOrder = [
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
    'Sunday',
  ];
  const targetHeaders: string[] = [
    'Cycle',
    'ACN',
    'First Name',
    'Last Name',
    'Address',
    'Suburb',
    'Post Code',
    'Contact',
    'Region',
    'Team',
    'Cover Team',
    'Order',
    'Fee',
    'Cancelations',
    'Notes',
    'Complaints/Incidents',
  ];

  // ---- Get Sheets ----
  const sourceSheet = workbook.getWorksheet(sourceSheetName);
  if (!sourceSheet) return;

  const existingSheet = workbook.getWorksheet(targetSheetName);
  if (existingSheet) existingSheet.delete();
  const newSheet = workbook.addWorksheet(targetSheetName);

  // ---- Read all source data at once ----
  const sourceData = sourceSheet.getUsedRange()?.getValues() as (
    | string
    | number
    | boolean
  )[][];
  if (!sourceData) return;
  const headers = sourceData[1] as string[];

  const cycleIndex = headers.indexOf('Cycle');
  const teamIndex = headers.indexOf('Team');
  const orderIndex = headers.indexOf('Order');
  if (cycleIndex < 0 || teamIndex < 0 || orderIndex < 0) return;

  // Map target headers to source indices
  const columnMapping: number[] = targetHeaders.map(h => headers.indexOf(h));

  // ---- Filter + Sort in memory ----
  const rows = sourceData
    .slice(2) // skip top 2 rows in Schedule
    .filter(r => (r[cycleIndex] ?? '').toString().includes('Week 1'))
    .map(r => {
      const day =
        (r[cycleIndex] ?? '').toString().match(/^(\w+)/)?.[1] ?? 'Unknown';
      const team = (r[teamIndex] ?? '').toString().trim() || 'Unassigned';
      const mapped = columnMapping.map(idx => (idx >= 0 ? r[idx] : ''));
      return { day, team, data: mapped };
    })
    .sort((a, b) => {
      const da = dayOrder.indexOf(a.day);
      const db = dayOrder.indexOf(b.day);
      return da - db || a.team.localeCompare(b.team);
    });

  // ---- Build output matrix ----
  const output: (string | number | boolean)[][] = [];
  output.push(targetHeaders);

  let lastDay = '',
    lastTeam = '';
  let groupRows: (string | number | boolean)[][] = [];

  const flushGroup = () => {
    // sort by Order column index in target headers
    const orderColIdx = targetHeaders.indexOf('Order');
    groupRows.sort((a, b) => {
      const aVal = a[orderColIdx];
      const bVal = b[orderColIdx];
      const aEmpty = aVal === '' || aVal === null;
      const bEmpty = bVal === '' || bVal === null;
      if (aEmpty && bEmpty) return 0;
      if (aEmpty) return 1; // blanks last
      if (bEmpty) return -1;
      return Number(aVal) - Number(bVal);
    });
    output.push(...groupRows);
    groupRows = [];
  };

  for (const row of rows) {
    if (row.day !== lastDay) {
      if (groupRows.length) flushGroup();
      output.push([
        `📅 ${row.day.toUpperCase()}`,
        ...Array(targetHeaders.length - 1).fill(''),
      ]);
      lastDay = row.day;
      lastTeam = '';
    }
    if (row.team !== lastTeam) {
      if (groupRows.length) flushGroup();
      output.push([
        `👥 ${row.team.toUpperCase()}`,
        ...Array(targetHeaders.length - 1).fill(''),
      ]);
      lastTeam = row.team;
    }
    groupRows.push(row.data);
  }
  if (groupRows.length) flushGroup();

  // ---- Write all data in one go ----
  const outRange = newSheet.getRangeByIndexes(
    0,
    0,
    output.length,
    targetHeaders.length
  );
  outRange.setValues(output);

  // ---- Bulk format headers ----
  const headerRow = newSheet.getRangeByIndexes(0, 0, 1, targetHeaders.length);
  headerRow
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRow.getFormat().getFont().setBold(true);
  headerRow.getFormat().setRowHeight(25);

  // ---- Bulk format day and team headers ----
  output.forEach((row, i) => {
    if (typeof row[0] === 'string' && (row[0] as string).startsWith('📅')) {
      const r = newSheet.getRangeByIndexes(i, 0, 1, targetHeaders.length);
      r.getFormat().getFill().setColor('#4A4A4A');
      r.getFormat().getFont().setColor('#FFFFFF');
      r.getFormat().getFont().setBold(true);
    }
    if (typeof row[0] === 'string' && (row[0] as string).startsWith('👥')) {
      const r = newSheet.getRangeByIndexes(i, 0, 1, targetHeaders.length);
      r.getFormat().getFill().setColor('#CCCCCC');
      r.getFormat().getFont().setBold(true);
    }
  });

  // ---- Auto fit once ----
  newSheet.getUsedRange()?.getFormat().autofitColumns();
  newSheet.getFreezePanes().freezeRows(1);
}
