# SuperQAT COM Add-in (Windows Only)

A native COM add-in for **Word, Excel, and PowerPoint** that can execute **any** ribbon command using `CommandBars.ExecuteMso()`. Unlike the web add-in which is limited by Office.js, this version can trigger every command — Shapes, Track Changes, Print dialogs, Copilot, everything.

## How it works

When you click a command in the panel, it calls:
```csharp
Application.CommandBars.ExecuteMso("Bold");  // or any MSO command ID
```
This is the same as clicking the actual ribbon button. No limitations.

## What you need

- **Windows 10/11**
- **Visual Studio 2022** (Community edition is free)
  - Workload: "Office/SharePoint development"
  - Workload: ".NET desktop development"
- **Microsoft Office** (Word, Excel, or PowerPoint)

## Build steps

1. Open `SuperQAT.sln` in Visual Studio 2022

2. If NuGet packages don't restore automatically:
   ```
   Tools > NuGet Package Manager > Package Manager Console
   Update-Package -Reinstall
   ```

3. Build the solution:
   ```
   Build > Build Solution (Ctrl+Shift+B)
   ```

4. Register the COM add-in (run CMD as Administrator):
   ```
   cd path\to\com-addin\SuperQAT\bin\Release
   regasm SuperQAT.dll /codebase /tlb
   ```

5. Open Word (or Excel/PowerPoint). You should see a **SuperQAT** tab in the ribbon.

6. Click **Open SuperQAT** to open the command panel.

## Using the panel

- **Click** any command to execute it instantly
- **Search** by typing in the search box (Ctrl+F to focus)
- **Enter** key executes the selected command
- **Escape** hides the panel
- The panel stays on top so you can keep using it while working

## Uninstall

Run CMD as Administrator:
```
regasm SuperQAT.dll /unregister
```

## Project structure

```
com-addin/
├── SuperQAT.sln              # Visual Studio solution
└── SuperQAT/
    ├── SuperQAT.csproj        # Project file (.NET Framework 4.8)
    ├── ThisAddIn.cs            # COM add-in entry point + ExecuteMso
    ├── CommandPanel.cs         # WinForms floating panel with search
    ├── MsoCommands.cs          # 726 MSO command IDs
    ├── SuperQATRibbon.xml      # Ribbon tab definition
    └── Properties/
        └── AssemblyInfo.cs
```

## Notes

- Some MSO command IDs only work when the right context is active (e.g. table commands only work when cursor is in a table)
- The add-in auto-registers for Word, Excel, and PowerPoint
- If a command isn't available in the current context, a message box explains why

## License

MIT
