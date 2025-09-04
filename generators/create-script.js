#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const scriptName = process.argv[2];
if (!scriptName) {
  console.error('❌ Usage: npm run new-script <script-name>');
  process.exit(1);
}

const templateContent = `/**
 * Office Script: ${scriptName}
 * Description: [Add description here]
 * Author: [Your name]
 * Created: ${new Date().toISOString().split('T')[0]}
 */

function main(workbook: ExcelScript.Workbook): void {
  // Get the active worksheet
  const worksheet = workbook.getActiveWorksheet();
  
  // TODO: Add your Office Script logic here
  console.log('Script "${scriptName}" is running...');
  
  // Example: Get a range and log its value
  // const range = worksheet.getRange('A1');
  // console.log('A1 value:', range.getValue());
}

// Helper functions (if needed)
// function helperFunction(): void {
//   // Add helper functions here
// }
`;

const filePath = path.join('src', 'scripts', `${scriptName}.ts`);
fs.writeFileSync(filePath, templateContent);

console.log(`✅ Created new Office Script: ${filePath}`);
console.log('💡 Tips:');
console.log('  - Use "npm run build" to check TypeScript compilation');
console.log('  - Use "npm run lint" to check code quality');
console.log('  - Copy the final code to Office Scripts when ready');
