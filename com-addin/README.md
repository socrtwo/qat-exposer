# SuperQAT COM Add-ins (VBA)

These COM-style VBA add-ins can **execute every single ribbon command** in
Word, Excel, and PowerPoint — including the ~2,000 commands that the web
add-in can only show guidance for.

They use `Application.CommandBars.ExecuteMso()` which directly triggers any
built-in ribbon button.

| App        | Commands | Add-in format |
|------------|----------|---------------|
| Word       | 1,334    | `.dotm`       |
| Excel      | 1,080    | `.xlam`       |
| PowerPoint | 769      | `.ppam`       |

## How to Install

### Word

1. Open Word on your desktop (Windows or Mac)
2. Press **Alt+F11** to open the VBA Editor
3. **File → Import File** — select `word/SuperQAT.bas`
4. **File → Import File** — select `word/SuperQATData.bas`
5. **File → Import File** — select `word/SuperQATForm.frm`
6. Close the VBA Editor
7. **File → Save As** — change type to **Word Macro-Enabled Template (.dotm)**
8. Save to your STARTUP folder:
   - Windows: `%appdata%\Microsoft\Word\STARTUP\SuperQAT.dotm`
   - Mac: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`
9. Restart Word — "SuperQAT" appears in the menu bar

### Excel

1. Open Excel, press **Alt+F11**
2. Import all three files from `excel/` folder (same as Word steps 3-5)
3. Close VBA Editor
4. **File → Save As** — change type to **Excel Add-In (.xlam)**
5. Save to:
   - Windows: `%appdata%\Microsoft\AddIns\SuperQAT.xlam`
   - Mac: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/`
6. **File → Options → Add-Ins → Manage: Excel Add-ins → Go**
7. Check **SuperQAT** and click OK

### PowerPoint

1. Open PowerPoint, press **Alt+F11**
2. Import all three files from `powerpoint/` folder
3. Close VBA Editor
4. **File → Save As** — change type to **PowerPoint Add-In (.ppam)**
5. Save to:
   - Windows: `%appdata%\Microsoft\AddIns\SuperQAT.ppam`
   - Mac: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/`
6. **File → Options → Add-Ins → Manage: PowerPoint Add-ins → Go**
7. Check **SuperQAT** and click OK

## How to Use

1. Click **SuperQAT** in the menu bar (or run the `ShowSuperQAT` macro)
2. A floating window appears with a search box and command list
3. Type to search — the list filters in real time
4. Select a command and click **Run Command** (or double-click it)
5. The command executes immediately — no ribbon navigation needed

## How It Works

Every command is executed via:

```vba
Application.CommandBars.ExecuteMso "Bold"
```

This is the same mechanism that the QAT (Quick Access Toolbar) uses
internally. It works for all built-in ribbon commands, including ones that
have no Office.js equivalent.

## Notes

- These are **desktop-only** (Windows and Mac). They do not work in
  Office for the web.
- Microsoft considers COM/VBA add-ins "legacy" but they still work in
  Office 365 desktop and will for the foreseeable future.
- Some commands may fail if they require a specific context (e.g., table
  commands need a table selected). The add-in shows an error message when
  this happens.
- The command data comes from Microsoft's official control ID files at
  [github.com/OfficeDev/office-fluent-ui-command-identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
