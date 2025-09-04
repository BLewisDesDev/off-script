# Office Scripts TypeScript Development Environment

This environment provides full TypeScript IntelliSense, linting, and formatting for Office Scripts development.

## Getting Started

1. **Write scripts** in `src/scripts/`
2. **Validate code** with `npm run check`
3. **Copy to Office** when ready

## Available Commands

- `npm run build` - Check TypeScript compilation
- `npm run lint` - Check code quality
- `npm run lint:fix` - Auto-fix linting issues
- `npm run format` - Format code with Prettier
- `npm run check` - Run both build and lint
- `npm run new-script <name>` - Create new script template

## Development Workflow

1. Create new script: `npm run new-script my-script`
2. Edit in `src/scripts/my-script.ts`
3. Validate: `npm run check`
4. Copy final code to Office Scripts platform

## Office Scripts Rules

- Always use `ExcelScript.Workbook` as first parameter
- Return `void` unless integrating with Power Automate
- Use `console.log()` for debugging (remove in production)
- Handle errors gracefully with proper checks
- Optimize for performance (get data once, process in memory)

## VS Code Features

- Full IntelliSense for Office Scripts APIs
- Real-time error detection
- Auto-formatting on save
- Organized imports
- Code snippets and suggestions

## API Reference

Your Office Scripts APIs are available through:

- `api/office_interface.d.ts`
- `api/office-scripts-types/`

## Tips

- Use TypeScript strict mode for better code quality
- Follow camelCase naming convention
- Write descriptive comments for complex logic
- Test scripts manually before automation
