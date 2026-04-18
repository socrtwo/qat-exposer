# SuperQAT — Word VBA Template

A Word template (.dotm) that adds a dropdown of all 1,884 QAT commands to your toolbar. Pick a command from the dropdown and it runs immediately on your document.

## Quick Start

### Step 1: Enable VBA project access (one-time setup)

1. Open **Word**
2. Go to **File > Options > Trust Center > Trust Center Settings**
3. Click **Macro Settings**
4. Check **Trust access to the VBA project object model**
5. Click **OK** and close Word

### Step 2: Build the template

Open PowerShell and run:

```powershell
cd word-vba-addin
powershell -ExecutionPolicy Bypass -File build.ps1
```

This creates **SuperQAT.dotm** in the same folder.

### Step 3: Install

**Option A — Load every time Word opens (recommended):**

Copy `SuperQAT.dotm` to your Word STARTUP folder:

```powershell
Copy-Item SuperQAT.dotm "$env:APPDATA\Microsoft\Word\STARTUP\"
```

**Option B — Load just this once:**

Double-click `SuperQAT.dotm` to open it in Word. The dropdown appears in the **Add-ins** tab.

## How to Use

1. Open any Word document
2. Click the **Add-ins** tab in the Ribbon
3. You'll see a **QAT Command** dropdown with all 1,884 commands
4. Type to search, or scroll to find a command
5. Select it — the command runs on your current document

## Files

| File | What it does |
|------|-------------|
| `commands.csv` | All 1,884 commands (ID + caption), extracted from the Excel source |
| `QATModule.bas` | VBA source code — the dropdown logic and command data |
| `build.ps1` | PowerShell script that assembles everything into SuperQAT.dotm |

## Uninstall

Delete `SuperQAT.dotm` from `%APPDATA%\Microsoft\Word\STARTUP\` and restart Word.
