# The Office Scripts Commandments

## Essential Rules for Error-Free Office Scripts Development

### **Priority Level 1: CRITICAL - Will Break Your Script**

#### **1. THOU SHALL NOT USE THE `any` TYPE**

- **NEVER** use `let value: any;` (explicit any)
- **NEVER** use `let value;` without initialization (implicit any)
- **ALWAYS** specify types or initialize with values:

  ```typescript
  // ❌ WRONG - Will cause compile error
  let value: any;
  let uninitialized;

  // ✅ CORRECT
  let value: string | number | boolean; // Union type
  let initialized = 5; // Inferred type
  let explicit: number; // Explicit type
  ```

#### **2. THOU SHALL ALWAYS HAVE A PROPER MAIN FUNCTION**

- **MUST** start with `function main(workbook: ExcelScript.Workbook)`
- **MUST** use `ExcelScript.Workbook` as first parameter type
- All executable code **MUST** be inside functions

  ```typescript
  // ✅ CORRECT - Required signature
  function main(workbook: ExcelScript.Workbook) {
    // Your code here
  }
  ```

#### **3. THOU SHALL NOT INHERIT FROM ExcelScript NAMESPACE**

- **NEVER** extend or implement any `ExcelScript` classes/interfaces
- **NEVER** create subclasses of Office Scripts objects
- Create your own interfaces separately from `ExcelScript` namespace

#### **4. THOU SHALL ONLY USE ARROW FUNCTIONS IN ARRAY CALLBACKS**

- **ONLY** use arrow functions `(x) => {}` in array methods
- **NEVER** use traditional `function(x) {}` syntax in callbacks

  ```typescript
  // ✅ CORRECT
  const filtered = myArray.filter(x => x > 5);

  // ❌ WRONG - Will cause compile error
  const filtered = myArray.filter(function (x) {
    return x > 5;
  });
  ```

#### **5. THOU SHALL NOT USE RESTRICTED IDENTIFIERS**

- **NEVER** use these reserved words as variable names:
  - `Excel`
  - `ExcelScript`
  - `console`

#### **6. THOU SHALL NOT USE FORBIDDEN FUNCTIONS**

- **NEVER** use `eval()` function (security restriction)
- **NEVER** use Office Scripts APIs in constructors
- **NEVER** use `console.log()` in constructors

### **Priority Level 2: PERFORMANCE - Will Slow Your Script**

#### **7. THOU SHALL MINIMIZE WORKBOOK COMMUNICATION**

- **Cache** range values in variables instead of repeated API calls
- **Read once, use many times** - store `getValues()` results
- **Batch operations** instead of individual cell updates

  ```typescript
  // ❌ SLOW - Multiple API calls
  for (let i = 0; i < 100; i++) {
    sheet.getRange(`A${i}`).setValue(data[i]);
  }

  // ✅ FAST - Single API call
  let values = data.map(d => [d]);
  sheet.getRange('A1:A100').setValues(values);
  ```

#### **8. THOU SHALL AVOID TRY-CATCH IN LOOPS**

- **NEVER** put `try-catch` blocks inside or around loops
- Use `try-catch` only for critical error handling
- Prefer validation checks over error catching

#### **9. THOU SHALL MANAGE CALCULATION MODE**

- Set calculation to manual for heavy operations
- Manually calculate when needed
- Restore automatic calculation afterward

  ```typescript
  const app = workbook.getApplication();
  app.setCalculationMode(ExcelScript.CalculationMode.manual);
  // ... perform operations
  app.calculate(ExcelScript.CalculationType.fullRebuild);
  app.setCalculationMode(ExcelScript.CalculationMode.automatic);
  ```

#### **10. THOU SHALL NOT LOOP OVER ROWS**

- **NEVER** iterate through rows individually
- **Get entire ranges** and work with 2D arrays
- **Set ranges** using arrays, not cell-by-cell

### **Priority Level 3: RELIABILITY - Will Make Scripts Robust**

#### **11. THOU SHALL ALWAYS VALIDATE OBJECT EXISTENCE**

- **Check** if worksheets, tables, ranges exist before using
- Use optional chaining `?.` or explicit `if` checks

  ```typescript
  // ✅ Safe approaches
  let sheet = workbook.getWorksheet('Data');
  if (sheet) {
    // Use sheet safely
  }

  // Or use optional chaining
  workbook.getWorksheet('Data')?.delete();
  ```

#### **12. THOU SHALL HANDLE ERRORS GRACEFULLY**

- Use `return` statements to exit cleanly
- Use `throw` only for Power Automate flows
- Provide meaningful error messages

  ```typescript
  if (!requiredTable) {
    console.log("Required table 'Sales' not found.");
    return; // Exit gracefully
  }
  ```

#### **13. THOU SHALL VALIDATE DATA TYPES**

- **Declare** expected types for cell values: `string | number | boolean`
- **Cast** values when needed: `value as string`
- **Check** data structure before processing

### **Priority Level 4: STRUCTURE - Code Organization**

#### **14. THOU SHALL USE PROPER ASYNC PATTERNS**

- Make `main` function `async` when using external calls
- Use `await` with `fetch()` operations
- Return `Promise<Type>` for async functions

  ```typescript
  async function main(workbook: ExcelScript.Workbook): Promise<void> {
    const response = await fetch('https://api.example.com/data');
    const data = await response.json();
  }
  ```

#### **15. THOU SHALL NOT USE BROWSER STORAGE**

- **NO** `localStorage` or `sessionStorage` (not supported)
- Store data in variables or pass through function parameters
- Use Power Automate for persistent data between runs

#### **16. THOU SHALL USE APPROPRIATE RETURN TYPES**

- Return `void` for scripts that don't return data
- Return specific types for Power Automate integration
- Use interfaces for complex return objects

### **Priority Level 5: BEST PRACTICES - Code Quality**

#### **17. THOU SHALL WRITE EFFICIENT LOOPS**

- **Get data once** before loops, not inside them
- **Process arrays** instead of individual cells
- **Minimize** API calls within iterations

#### **18. THOU SHALL USE MEANINGFUL NAMES**

- Use descriptive variable names
- Follow camelCase convention
- Avoid abbreviations unless widely understood

#### **19. THOU SHALL HANDLE EXTERNAL API LIMITATIONS**

- No OAuth2 or sign-in flows available
- Hardcode API keys (security limitation)
- Handle fetch errors appropriately
- Check API response structure

#### **20. THOU SHALL OPTIMIZE FOR POWER AUTOMATE**

- Remove all `console.log()` statements for production
- Use `throw` to stop flows on critical errors
- Return data in JSON-compatible formats
- Test scripts manually before automation

### **Priority Level 6: DEBUGGING - Development Aids**

#### **21. THOU SHALL USE ACTION RECORDER WISELY**

- Use Action Recorder to learn API patterns
- **Always** review and optimize generated code
- Remove unnecessary selections and activations

#### **22. THOU SHALL LOG STRATEGICALLY**

- Use `console.log()` for debugging during development
- Remove logging from production scripts
- Log meaningful information, not just variable dumps

#### **23. THOU SHALL HANDLE PLATFORM DIFFERENCES**

- Some APIs only work in Excel on the web
- Test scripts on target platforms
- Provide fallbacks for unsupported features

### **Quick Reference: Common Patterns**

#### **Safe Object Access**

```typescript
// Check existence first
let table = workbook.getTable('MyTable');
if (table) {
  // Use table safely
}

// Or use optional chaining
workbook.getTable('MyTable')?.delete();
```

#### **Efficient Data Handling**

```typescript
// Get data once, process in memory
let range = sheet.getRange('A1:Z100');
let values = range.getValues();

// Process the 2D array
for (let row = 0; row < values.length; row++) {
  for (let col = 0; col < values[row].length; col++) {
    // Process values[row][col]
  }
}

// Write back if needed
range.setValues(modifiedValues);
```

#### **Proper Type Declarations**

```typescript
// For cell values that can be multiple types
let cellValue: string | number | boolean = range.getValue();

// For known single types
let textValue = range.getValue() as string;

// For arrays
let tableData: (string | number | boolean)[][] = table
  .getRangeBetweenHeaderAndTotal()
  .getValues();
```

#### **Error Handling Template**

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Validate inputs first
  const sheet = workbook.getWorksheet('Data');
  if (!sheet) {
    console.log("Worksheet 'Data' not found.");
    return;
  }

  const table = sheet.getTable('MyTable');
  if (!table) {
    console.log("Table 'MyTable' not found.");
    return;
  }

  // Proceed with operations
  try {
    // Risky operations here
  } catch (error) {
    console.log(`Error occurred: ${error}`);
    return;
  }
}
```

---

## Remember: These commandments are ordered by criticality. Focus on Priority 1 rules first to avoid compilation errors, then work down the priority levels to improve performance and reliability.
