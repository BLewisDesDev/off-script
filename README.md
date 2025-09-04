# Office Scripts TypeScript Development Environment

> A comprehensive development environment for creating, testing, and managing Microsoft Office Scripts with full TypeScript support, IntelliSense, and modern development tooling.

## 🚀 Quick Start

```bash
# Clone and setup
git clone <repository-url>
cd office_script_environment
npm install

# Create your first script
npm run new-script my-first-script

# Validate your code
npm run check
```

## 📋 What This Project Provides

This environment eliminates the pain points of Office Scripts development by providing:

- **Full TypeScript IntelliSense** - Complete Office Scripts API type definitions
- **Modern Development Workflow** - Linting, formatting, and error checking
- **Script Templates** - Pre-built templates for common Office Scripts patterns
- **Production Scripts** - Real-world examples for business operations
- **Best Practices Guide** - Comprehensive rules and patterns for reliable scripts
- **Development Tools** - Script generators and utilities

## 🛠️ Development Workflow

### Creating New Scripts

```bash
# Generate a new script with template
npm run new-script data-processor

# Edit in scripts/templates/data-processor.ts
# Full IntelliSense and error checking available
```

### Validation & Quality

```bash
npm run build      # Check TypeScript compilation
npm run lint       # Check code quality with ESLint
npm run lint:fix   # Auto-fix linting issues
npm run format     # Format code with Prettier
npm run check      # Run build + lint together
```

### Script Organization

```
scripts/
├── templates/           # Example scripts and templates
├── utilities/          # Utility functions and helpers
├── hm_ops/            # Human resources operations
├── acp_sc_reconciliation/  # Accounting reconciliation
└── [your-domain]/     # Your business-specific scripts
```

## 📚 Key Features

### 🔧 Complete Office Scripts API Support

The `api/office_interface.d.ts` file provides comprehensive TypeScript definitions for:

- **Workbook Operations** - Sheets, ranges, tables, charts
- **Data Manipulation** - Reading, writing, formatting
- **Advanced Features** - Pivot tables, slicers, queries
- **Error Handling** - Proper type checking and validation

### 📖 Comprehensive Documentation

- **Office Scripts Commandments** (`docs/office_scripts_commandments.md`) - 23 essential rules for error-free development
- **Development Guide** (`docs/DEVELOPMENT.md`) - Workflow and best practices
- **Setup Guide** (`docs/simple_office_scripts_setup.md`) - Complete setup instructions for Mac
- **Function Reference** (`docs/office_scripts_function_list.md`) - Complete API documentation

### 🏗️ Production-Ready Examples

**Data Processing:**

- `scripts/utilities/strip_date.ts` - Date format standardization across workbooks
- `scripts/templates/example-highlight-cells.ts` - Conditional formatting based on values

**Business Operations:**

- `scripts/hm_ops/get_regions.ts` - Postcode-to-region lookups
- `scripts/hm_ops/import_cwk2.ts` - Data import operations
- `scripts/acp_sc_reconciliation/` - Accounting reconciliation workflows

## 🎯 Office Scripts Best Practices

This environment enforces critical Office Scripts patterns:

### ✅ Required Patterns

```typescript
// Always use proper main function signature
function main(workbook: ExcelScript.Workbook): void {
  // Your code here
}

// Proper type declarations for cell values
let cellValue: string | number | boolean = range.getValue();

// Efficient batch operations
const values = range.getValues(); // Get once
// Process in memory
range.setValues(modifiedValues); // Set once
```

### ❌ Avoid These Mistakes

```typescript
// NEVER use any type
let value: any; // ❌ Will cause compile errors

// NEVER loop over individual cells
for (let i = 1; i <= 1000; i++) {
  sheet.getRange(`A${i}`).setValue(data[i]); // ❌ Slow
}

// NEVER use function declarations in callbacks
array.filter(function (x) {
  return x > 5;
}); // ❌ Compilation error
```

## 🔧 Configuration

### TypeScript Configuration

The `tsconfig.json` is optimized for Office Scripts:

- **Target**: ES2017 (Office Scripts runtime)
- **Strict Mode**: Enabled for better error catching
- **Module System**: ES2022 with Node resolution
- **Type Roots**: Includes Office Scripts API definitions

### ESLint & Prettier

- **ESLint**: Configured with TypeScript rules and Prettier integration
- **Prettier**: Consistent code formatting
- **VS Code Integration**: Auto-format on save

## 📦 Dependencies

### Development Dependencies

- **TypeScript** - Language support and compilation
- **ESLint** - Code quality and error detection
- **Prettier** - Code formatting
- **@types/node** - Node.js type definitions

### Office Scripts API

- Custom type definitions in `api/office_interface.d.ts`
- Comprehensive coverage of Excel Script APIs
- Regular updates to match Microsoft's Office Scripts platform

## 🚦 Project Scripts

| Command                     | Purpose                                     |
| --------------------------- | ------------------------------------------- |
| `npm run build`             | Check TypeScript compilation without output |
| `npm run dev`               | Watch mode for continuous validation        |
| `npm run lint`              | Run ESLint on all TypeScript files          |
| `npm run lint:fix`          | Auto-fix ESLint issues                      |
| `npm run format`            | Format code with Prettier                   |
| `npm run check`             | Run build + lint together                   |
| `npm run clean`             | Remove dist directory                       |
| `npm run new-script <name>` | Create new script from template             |

## 🎓 Learning Resources

### Essential Reading

1. **Start Here**: `docs/DEVELOPMENT.md` - Basic workflow
2. **Critical Rules**: `docs/office_scripts_commandments.md` - Avoid common pitfalls
3. **Complete Setup**: `docs/simple_office_scripts_setup.md` - Mac development setup
4. **API Reference**: `docs/office_scripts_function_list.md` - Function documentation

### Example Scripts

- **Beginner**: `scripts/templates/example-highlight-cells.ts`
- **Intermediate**: `scripts/utilities/strip_date.ts`
- **Advanced**: `scripts/hm_ops/get_regions.ts`

## 🤝 Contributing

### Adding New Scripts

1. Use `npm run new-script <name>` to generate template
2. Follow the Office Scripts Commandments (Priority 1 rules are critical)
3. Run `npm run check` before finalizing
4. Add documentation comments for complex logic

### Code Standards

- **TypeScript**: Strict mode enabled, no `any` types
- **Error Handling**: Graceful exits with console logging
- **Performance**: Batch operations, minimal API calls
- **Naming**: camelCase, descriptive variable names

## 🐛 Troubleshooting

### Common Issues

**"Cannot find module" errors:**

- Ensure `npm install` was run
- Check `typeRoots` in `tsconfig.json`

**Compilation errors:**

- Review Office Scripts Commandments (Priority 1)
- Avoid `any` types and uninitialized variables
- Use proper main function signature

**Performance issues:**

- Follow batch operation patterns
- Minimize workbook API calls
- Process data in memory, not cell-by-cell

### Getting Help

1. Check the Office Scripts Commandments for rule violations
2. Review example scripts for patterns
3. Validate with `npm run check` for immediate feedback

## 📄 License

This project provides development tooling and examples for Microsoft Office Scripts development. Office Scripts itself is a Microsoft product with its own licensing terms.

---

**Ready to build powerful Office Scripts?** Start with `npm run new-script your-script-name` and follow the development workflow!
