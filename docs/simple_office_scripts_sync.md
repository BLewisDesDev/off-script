# Simple Office Scripts Development Setup for Mac

Stop copy-pasting! This guide shows you how to write Office Scripts in VS Code with TypeScript and automatically sync them to Excel.

## Quick Start

### 1. Find Your OneDrive Office Scripts Folder

First, locate where your Office Scripts are stored. Check these locations:

```bash
# Most common locations on Mac
ls "$HOME/OneDrive - Personal/Office Scripts"
ls "$HOME/OneDrive/Office Scripts"
ls "$HOME/Library/CloudStorage/OneDrive-Personal/Office Scripts"
```

Note the path that exists - you'll need it in step 3.

### 2. Create Your Project Structure

```bash
mkdir office-scripts-dev
cd office-scripts-dev
npm init -y
mkdir src
```

### 3. Create the Deploy Script

Create `deploy.js` in your project root:

```javascript
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// UPDATE THIS PATH to match your OneDrive location from step 1
const ONEDRIVE_SCRIPTS_PATH = path.join(
  process.env.HOME,
  'OneDrive - Personal',
  'Office Scripts'
);

function deployScript(sourceFilePath, shouldSwitchToExcel = false) {
  if (!fs.existsSync(sourceFilePath)) {
    console.error('❌ Source file not found:', sourceFilePath);
    return;
  }

  if (!fs.existsSync(ONEDRIVE_SCRIPTS_PATH)) {
    console.error(
      '❌ OneDrive Office Scripts folder not found at:',
      ONEDRIVE_SCRIPTS_PATH
    );
    console.error('Please update ONEDRIVE_SCRIPTS_PATH in deploy.js');
    return;
  }

  // Read your TypeScript source code
  const sourceContent = fs.readFileSync(sourceFilePath, 'utf8');
  const fileName = path.basename(sourceFilePath, '.ts');
  const ostsFileName = `${fileName}.osts`;
  const ostsPath = path.join(ONEDRIVE_SCRIPTS_PATH, ostsFileName);

  let ostsContent;

  if (fs.existsSync(ostsPath)) {
    // Update existing .osts file
    try {
      ostsContent = JSON.parse(fs.readFileSync(ostsPath, 'utf8'));
      ostsContent.body = sourceContent;
      ostsContent.lastModified = new Date().toISOString();
    } catch (error) {
      console.error('❌ Error reading existing .osts file:', error.message);
      return;
    }
  } else {
    // Create new .osts file
    ostsContent = {
      version: '0.3.0',
      body: sourceContent,
      metadata: {
        name: fileName,
        description: '',
        created: new Date().toISOString(),
        lastModified: new Date().toISOString(),
      },
    };
  }

  try {
    fs.writeFileSync(ostsPath, JSON.stringify(ostsContent, null, 2));
    console.log(`✅ Deployed ${fileName} to Office Scripts`);

    if (shouldSwitchToExcel) {
      try {
        execSync(
          'osascript -e \'tell application "Microsoft Excel" to activate\'',
          { stdio: 'ignore' }
        );
        console.log('🔄 Switched to Excel');
      } catch (error) {
        console.log('⚠️  Could not switch to Excel (not running?)');
      }
    }
  } catch (error) {
    console.error('❌ Error writing .osts file:', error.message);
  }
}

// Command line usage
const sourceFile = process.argv[2];
const shouldSwitch = process.argv.includes('--switch');

if (sourceFile) {
  deployScript(sourceFile, shouldSwitch);
} else {
  console.error('❌ Usage: node deploy.js path/to/script.ts [--switch]');
  console.error('Example: node deploy.js src/my-script.ts --switch');
}
```

### 4. Set Up VS Code

Create `.vscode/tasks.json` in your project:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Deploy Office Script",
      "type": "shell",
      "command": "node",
      "args": ["deploy.js", "${file}", "--switch"],
      "group": "build",
      "presentation": {
        "echo": true,
        "reveal": "silent",
        "focus": false,
        "panel": "shared"
      },
      "problemMatcher": []
    }
  ]
}
```

Create `.vscode/keybindings.json` (or add to existing):

```json
[
  {
    "key": "cmd+shift+d",
    "command": "workbench.action.tasks.runTask",
    "args": "Deploy Office Script"
  }
]
```

### 5. Write Your Office Script

Create `src/my-script.ts`:

```typescript
function main(workbook: ExcelScript.Workbook): void {
  const worksheet = workbook.getActiveWorksheet();
  const range = worksheet.getRange('A1');
  range.setValue('Hello from VS Code!');

  console.log('Script executed successfully!');
}
```

## Usage

### Method 1: Keyboard Shortcut (Recommended)

1. Open your TypeScript file in VS Code
2. Press **Cmd+Shift+D**
3. Your script is deployed and Excel opens automatically!

### Method 2: Command Line

```bash
# Deploy and switch to Excel
node deploy.js src/my-script.ts --switch

# Deploy only (no window switching)
node deploy.js src/my-script.ts
```

### Method 3: VS Code Command Palette

1. Press **Cmd+Shift+P**
2. Type "Tasks: Run Task"
3. Select "Deploy Office Script"

## How It Works

1. **You write** TypeScript in `src/my-script.ts` with full IntelliSense
2. **Script reads** your TypeScript source code
3. **Script updates** the corresponding `.osts` file in OneDrive
4. **OneDrive syncs** the changes to Microsoft's servers
5. **Excel refreshes** the Office Scripts list automatically
6. **You run** the updated script in Excel - no copy/paste needed!

## The .osts File Format

Office Scripts are stored as JSON files with this structure:

```json
{
  "version": "0.3.0",
  "body": "function main(workbook: ExcelScript.Workbook): void {\n  // Your TypeScript code here\n}",
  "metadata": {
    "name": "my-script",
    "description": "",
    "created": "2025-09-03T10:30:00.000Z",
    "lastModified": "2025-09-03T10:35:00.000Z"
  }
}
```

The deploy script simply updates the `body` field with your TypeScript source code.

## Tips

- **Keep Excel open** while developing for faster script refresh
- **Name your files descriptively** - the .ts filename becomes the Office Script name
- **Use console.log()** for debugging - output appears in Excel's console
- **OneDrive must be running** for sync to work
- **Wait a few seconds** after deployment before running the script in Excel

## Troubleshooting

**Script not appearing in Excel?**

- Check that OneDrive is syncing (look for the cloud icon in menu bar)
- Verify the ONEDRIVE_SCRIPTS_PATH in deploy.js is correct
- Wait 10-15 seconds and refresh the Office Scripts panel in Excel

**"OneDrive folder not found" error?**

- Update the ONEDRIVE_SCRIPTS_PATH in deploy.js with the correct path from step 1

**Excel not switching focus?**

- Make sure Excel is running before using --switch
- The script will still deploy successfully even if Excel switching fails

## Next Steps

Once you're comfortable with this setup, you might want to:

- Set up a file watcher for automatic deployment on save
- Add TypeScript compilation and error checking
- Create templates for common Office Scripts patterns

But for now, you have a simple, reliable workflow that eliminates copy-pasting!
