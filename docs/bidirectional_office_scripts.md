# Bidirectional Office Scripts Sync for Mac

Sync Office Scripts both ways: VS Code ↔ OneDrive. Write in VS Code OR create scripts in Excel - they stay synchronized automatically.

## Setup

### 1. Enhanced Deploy Script

Replace your `deploy.js` with this bidirectional version:

```javascript
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// UPDATE THIS PATH to match your OneDrive location
const ONEDRIVE_SCRIPTS_PATH = path.join(
  process.env.HOME,
  'OneDrive - Personal',
  'Office Scripts'
);
const SOURCE_DIR = './src';

// Ensure source directory exists
if (!fs.existsSync(SOURCE_DIR)) {
  fs.mkdirSync(SOURCE_DIR, { recursive: true });
}

class OfficeScriptsSync {
  // Deploy .ts file to .osts (VS Code → OneDrive)
  static deployToOneDrive(sourceFilePath, shouldSwitchToExcel = false) {
    if (!fs.existsSync(sourceFilePath)) {
      console.error('❌ Source file not found:', sourceFilePath);
      return;
    }

    if (!fs.existsSync(ONEDRIVE_SCRIPTS_PATH)) {
      console.error(
        '❌ OneDrive Office Scripts folder not found at:',
        ONEDRIVE_SCRIPTS_PATH
      );
      return;
    }

    const sourceContent = fs.readFileSync(sourceFilePath, 'utf8');
    const fileName = path.basename(sourceFilePath, '.ts');
    const ostsPath = path.join(ONEDRIVE_SCRIPTS_PATH, `${fileName}.osts`);

    let ostsContent;

    if (fs.existsSync(ostsPath)) {
      try {
        ostsContent = JSON.parse(fs.readFileSync(ostsPath, 'utf8'));
        ostsContent.body = sourceContent;
        ostsContent.lastModified = new Date().toISOString();
      } catch (error) {
        console.error('❌ Error reading existing .osts file:', error.message);
        return;
      }
    } else {
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
      console.log(`✅ Deployed ${fileName}.ts → ${fileName}.osts`);

      if (shouldSwitchToExcel) {
        this.switchToExcel();
      }
    } catch (error) {
      console.error('❌ Error writing .osts file:', error.message);
    }
  }

  // Import .osts file to .ts (OneDrive → VS Code)
  static importFromOneDrive(ostsFileName) {
    const ostsPath = path.join(ONEDRIVE_SCRIPTS_PATH, ostsFileName);

    if (!fs.existsSync(ostsPath)) {
      console.error('❌ .osts file not found:', ostsPath);
      return;
    }

    try {
      const ostsContent = JSON.parse(fs.readFileSync(ostsPath, 'utf8'));
      const typeScriptCode = ostsContent.body;

      if (!typeScriptCode) {
        console.error('❌ No TypeScript code found in .osts file');
        return;
      }

      const baseName = path.basename(ostsFileName, '.osts');
      const tsPath = path.join(SOURCE_DIR, `${baseName}.ts`);

      // Check if local file exists and is different
      if (fs.existsSync(tsPath)) {
        const existingContent = fs.readFileSync(tsPath, 'utf8');
        if (existingContent === typeScriptCode) {
          console.log(`📋 ${baseName}.ts is already up to date`);
          return;
        }
        console.log(`🔄 Updating existing ${baseName}.ts`);
      } else {
        console.log(`✨ Creating new ${baseName}.ts`);
      }

      fs.writeFileSync(tsPath, typeScriptCode);
      console.log(`✅ Imported ${ostsFileName} → ${baseName}.ts`);
    } catch (error) {
      console.error('❌ Error importing .osts file:', error.message);
    }
  }

  // Sync all .osts files to .ts files (full import)
  static syncAllFromOneDrive(force = false) {
    if (!fs.existsSync(ONEDRIVE_SCRIPTS_PATH)) {
      console.error('❌ OneDrive Office Scripts folder not found');
      return;
    }

    console.log('📥 Syncing all Office Scripts from OneDrive...');

    const ostsFiles = fs
      .readdirSync(ONEDRIVE_SCRIPTS_PATH)
      .filter(file => file.endsWith('.osts'));

    if (ostsFiles.length === 0) {
      console.log('📭 No .osts files found in OneDrive');
      return;
    }

    console.log(`Found ${ostsFiles.length} Office Scripts to import:`);
    ostsFiles.forEach(file => console.log(`  - ${file}`));
    console.log('');

    let imported = 0;
    let skipped = 0;
    let errors = 0;

    ostsFiles.forEach(ostsFile => {
      try {
        const ostsPath = path.join(ONEDRIVE_SCRIPTS_PATH, ostsFile);
        const ostsContent = JSON.parse(fs.readFileSync(ostsPath, 'utf8'));
        const typeScriptCode = ostsContent.body;

        if (!typeScriptCode) {
          console.log(`⚠️  Skipped ${ostsFile} (no TypeScript code)`);
          skipped++;
          return;
        }

        const baseName = path.basename(ostsFile, '.osts');
        const tsPath = path.join(SOURCE_DIR, `${baseName}.ts`);

        // Check if we should update
        let shouldUpdate = force;
        if (!shouldUpdate) {
          if (!fs.existsSync(tsPath)) {
            shouldUpdate = true;
          } else {
            const existingContent = fs.readFileSync(tsPath, 'utf8');
            shouldUpdate = existingContent !== typeScriptCode;
          }
        }

        if (shouldUpdate) {
          fs.writeFileSync(tsPath, typeScriptCode);
          console.log(`✅ ${ostsFile} → ${baseName}.ts`);
          imported++;
        } else {
          console.log(`📋 ${baseName}.ts (already up to date)`);
          skipped++;
        }
      } catch (error) {
        console.error(`❌ Error processing ${ostsFile}:`, error.message);
        errors++;
      }
    });

    console.log('');
    console.log(
      `📊 Import Summary: ${imported} imported, ${skipped} skipped, ${errors} errors`
    );
  }

  // List all Office Scripts in OneDrive
  static listOneDriveScripts() {
    if (!fs.existsSync(ONEDRIVE_SCRIPTS_PATH)) {
      console.error('❌ OneDrive Office Scripts folder not found');
      return;
    }

    const ostsFiles = fs
      .readdirSync(ONEDRIVE_SCRIPTS_PATH)
      .filter(file => file.endsWith('.osts'));

    if (ostsFiles.length === 0) {
      console.log('📭 No Office Scripts found in OneDrive');
      return;
    }

    console.log(`📋 Found ${ostsFiles.length} Office Scripts in OneDrive:`);
    console.log('');

    ostsFiles.forEach(ostsFile => {
      try {
        const ostsPath = path.join(ONEDRIVE_SCRIPTS_PATH, ostsFile);
        const stats = fs.statSync(ostsPath);
        const ostsContent = JSON.parse(fs.readFileSync(ostsPath, 'utf8'));

        const baseName = path.basename(ostsFile, '.osts');
        const tsPath = path.join(SOURCE_DIR, `${baseName}.ts`);
        const hasLocal = fs.existsSync(tsPath);

        console.log(`📄 ${ostsFile}`);
        console.log(`   Modified: ${stats.mtime.toLocaleString()}`);
        console.log(`   Local .ts: ${hasLocal ? '✅ exists' : '❌ missing'}`);
        console.log(
          `   Size: ${ostsContent.body ? ostsContent.body.length : 0} characters`
        );
        console.log('');
      } catch (error) {
        console.log(`📄 ${ostsFile} (⚠️  parsing error)`);
        console.log('');
      }
    });
  }

  static switchToExcel() {
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
}

// Command line interface
const command = process.argv[2];

switch (command) {
  case 'deploy':
    const sourceFile = process.argv[3];
    const shouldSwitch = process.argv.includes('--switch');
    if (sourceFile) {
      OfficeScriptsSync.deployToOneDrive(sourceFile, shouldSwitch);
    } else {
      console.error(
        '❌ Usage: node deploy.js deploy path/to/script.ts [--switch]'
      );
    }
    break;

  case 'import':
    const ostsFile = process.argv[3];
    if (ostsFile) {
      OfficeScriptsSync.importFromOneDrive(ostsFile);
    } else {
      console.error('❌ Usage: node deploy.js import script-name.osts');
    }
    break;

  case 'sync':
    const force = process.argv.includes('--force');
    OfficeScriptsSync.syncAllFromOneDrive(force);
    break;

  case 'list':
    OfficeScriptsSync.listOneDriveScripts();
    break;

  default:
    console.log('Office Scripts Bidirectional Sync');
    console.log('');
    console.log('Commands:');
    console.log('  deploy <file.ts> [--switch]  Deploy TypeScript to OneDrive');
    console.log(
      '  import <file.osts>           Import single Office Script to TypeScript'
    );
    console.log(
      '  sync [--force]               Import all Office Scripts from OneDrive'
    );
    console.log(
      '  list                         List all Office Scripts in OneDrive'
    );
    console.log('');
    console.log('Examples:');
    console.log('  node deploy.js deploy src/my-script.ts --switch');
    console.log('  node deploy.js import my-script.osts');
    console.log('  node deploy.js sync');
    console.log('  node deploy.js list');
}
```

### 2. File Watcher for Automatic Sync

Create `watcher.js` for automatic bidirectional syncing:

```javascript
const fs = require('fs');
const path = require('path');
const chokidar = require('chokidar'); // npm install chokidar

// UPDATE THIS PATH
const ONEDRIVE_SCRIPTS_PATH = path.join(
  process.env.HOME,
  'OneDrive - Personal',
  'Office Scripts'
);
const SOURCE_DIR = './src';

class BidirectionalWatcher {
  constructor() {
    this.isRunning = false;
    this.tsWatcher = null;
    this.ostsWatcher = null;
    this.processing = new Set(); // Prevent infinite loops
  }

  log(message) {
    const timestamp = new Date().toLocaleTimeString();
    console.log(`[${timestamp}] ${message}`);
  }

  async processTypeScriptChange(filePath) {
    const fileName = path.basename(filePath);
    if (this.processing.has(fileName)) return;

    try {
      this.processing.add(fileName);
      this.log(`📝 ${fileName} changed, syncing to OneDrive...`);

      const { OfficeScriptsSync } = require('./deploy.js');
      OfficeScriptsSync.deployToOneDrive(filePath, false);
    } catch (error) {
      this.log(`❌ Error syncing ${fileName}: ${error.message}`);
    } finally {
      setTimeout(() => this.processing.delete(fileName), 2000);
    }
  }

  async processOstsChange(filePath) {
    const fileName = path.basename(filePath);
    const baseName = path.basename(fileName, '.osts');
    const tsFileName = `${baseName}.ts`;

    if (this.processing.has(tsFileName)) return;

    try {
      this.processing.add(tsFileName);
      this.log(`📥 ${fileName} changed in OneDrive, syncing to VS Code...`);

      const { OfficeScriptsSync } = require('./deploy.js');
      OfficeScriptsSync.importFromOneDrive(fileName);
    } catch (error) {
      this.log(`❌ Error importing ${fileName}: ${error.message}`);
    } finally {
      setTimeout(() => this.processing.delete(tsFileName), 2000);
    }
  }

  start() {
    if (this.isRunning) {
      this.log('⚠️  Watcher is already running');
      return;
    }

    if (!fs.existsSync(ONEDRIVE_SCRIPTS_PATH)) {
      this.log('❌ OneDrive Office Scripts folder not found');
      return;
    }

    if (!fs.existsSync(SOURCE_DIR)) {
      fs.mkdirSync(SOURCE_DIR, { recursive: true });
    }

    this.log('🚀 Starting bidirectional Office Scripts sync...');
    this.log(`📁 TypeScript source: ${SOURCE_DIR}`);
    this.log(`☁️  OneDrive scripts: ${ONEDRIVE_SCRIPTS_PATH}`);

    // Watch TypeScript files
    this.tsWatcher = chokidar.watch(`${SOURCE_DIR}/**/*.ts`, {
      ignored: /node_modules/,
      persistent: true,
    });

    this.tsWatcher.on('change', filePath => {
      this.processTypeScriptChange(filePath);
    });

    // Watch .osts files
    this.ostsWatcher = chokidar.watch(`${ONEDRIVE_SCRIPTS_PATH}/**/*.osts`, {
      ignored: /node_modules/,
      persistent: true,
    });

    this.ostsWatcher.on('change', filePath => {
      this.processOstsChange(filePath);
    });

    this.ostsWatcher.on('add', filePath => {
      // New .osts file created in OneDrive
      setTimeout(() => this.processOstsChange(filePath), 1000);
    });

    this.isRunning = true;
    this.log('👀 Watching for changes... (Press Ctrl+C to stop)');
  }

  stop() {
    if (this.tsWatcher) this.tsWatcher.close();
    if (this.ostsWatcher) this.ostsWatcher.close();
    this.isRunning = false;
    this.log('👋 Stopped watching');
  }
}

// CLI usage
if (require.main === module) {
  const watcher = new BidirectionalWatcher();

  process.on('SIGINT', () => {
    console.log('\\n');
    watcher.stop();
    process.exit(0);
  });

  watcher.start();
}

module.exports = BidirectionalWatcher;
```

### 3. Updated VS Code Tasks

Replace your `.vscode/tasks.json`:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Deploy to OneDrive",
      "type": "shell",
      "command": "node",
      "args": ["deploy.js", "deploy", "${file}", "--switch"],
      "group": "build",
      "presentation": {
        "echo": true,
        "reveal": "silent"
      }
    },
    {
      "label": "Sync All from OneDrive",
      "type": "shell",
      "command": "node",
      "args": ["deploy.js", "sync"],
      "group": "build",
      "presentation": {
        "echo": true,
        "reveal": "always"
      }
    },
    {
      "label": "List OneDrive Scripts",
      "type": "shell",
      "command": "node",
      "args": ["deploy.js", "list"],
      "group": "build",
      "presentation": {
        "echo": true,
        "reveal": "always"
      }
    },
    {
      "label": "Start Bidirectional Sync",
      "type": "shell",
      "command": "node",
      "args": ["watcher.js"],
      "group": "build",
      "presentation": {
        "echo": true,
        "reveal": "always",
        "panel": "new"
      },
      "isBackground": true,
      "runOptions": {
        "instanceLimit": 1
      }
    }
  ]
}
```

### 4. Updated Keyboard Shortcuts

Add to `.vscode/keybindings.json`:

```json
[
  {
    "key": "cmd+shift+d",
    "command": "workbench.action.tasks.runTask",
    "args": "Deploy to OneDrive"
  },
  {
    "key": "cmd+shift+s",
    "command": "workbench.action.tasks.runTask",
    "args": "Sync All from OneDrive"
  },
  {
    "key": "cmd+shift+l",
    "command": "workbench.action.tasks.runTask",
    "args": "List OneDrive Scripts"
  }
]
```

## Usage

### Manual Commands

```bash
# Deploy TypeScript to OneDrive (VS Code → Excel)
node deploy.js deploy src/my-script.ts --switch

# Import single Office Script (Excel → VS Code)
node deploy.js import my-script.osts

# Sync ALL Office Scripts from OneDrive to VS Code
node deploy.js sync

# List all Office Scripts in OneDrive
node deploy.js list

# Force sync (overwrite local files)
node deploy.js sync --force
```

### Automatic Bidirectional Sync

```bash
# Install file watcher dependency
npm install chokidar

# Start automatic bidirectional sync
node watcher.js
```

When the watcher is running:

- **Save a .ts file in VS Code** → automatically deployed to OneDrive
- **Create/edit script in Excel** → automatically imported to VS Code
- Works in both directions simultaneously!

### VS Code Shortcuts

- **Cmd+Shift+D**: Deploy current file to OneDrive
- **Cmd+Shift+S**: Sync all scripts from OneDrive
- **Cmd+Shift+L**: List all OneDrive scripts

## Workflow Examples

### Scenario 1: You create a script in Excel

1. Write an Office Script directly in Excel's editor
2. Save it in Excel
3. The watcher automatically creates `script-name.ts` in your `src/` folder
4. Continue editing in VS Code with full TypeScript support

### Scenario 2: Colleague shares a script

1. They share an Office Script via OneDrive/Teams
2. Run `node deploy.js sync` to pull all shared scripts
3. All scripts appear as `.ts` files in your `src/` folder
4. Edit and deploy back using your normal workflow

### Scenario 3: Full bidirectional development

1. Start the watcher: `node watcher.js`
2. Edit files in VS Code OR Excel - changes sync automatically
3. Multiple developers can collaborate on the same scripts
4. All changes are preserved in both locations

## Benefits

✅ **Write anywhere**: VS Code for complex development, Excel for quick fixes
✅ **Never lose work**: Scripts exist in both locations
✅ **Team collaboration**: Share via OneDrive, everyone gets TypeScript files
✅ **Version control**: Commit `.ts` files to Git
✅ **IntelliSense**: Full TypeScript support in VS Code
✅ **Automatic sync**: No manual copy-pasting ever again

This setup gives you the best of both worlds - the power of VS Code for development and the convenience of Excel's built-in Office Scripts editor!
